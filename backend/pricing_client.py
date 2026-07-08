"""
pricing_client.py — Agent-based AWS Pricing Estimation

The Bedrock agent autonomously:
1. Reads the MOM and decides which AWS services are needed
2. Calls MCP tools (search_services, get_service_fields, create_estimate,
   add_service, validate_estimate, export_estimate) in the right order
3. Iterates on errors (field grounding, lint failures) until it gets a valid URL

No hardcoded service lists, no hardcoded configs, no manual tool orchestration.
"""

import json
import re
import subprocess
import threading
import time
import os
import shutil
import boto3
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))


# ─── MCP stdio transport ──────────────────────────────────────────────────────

def _find_npx() -> str:
    found = shutil.which("npx")
    if found:
        return found
    candidate = r"C:\Program Files\nodejs\npx.cmd"
    if os.path.exists(candidate):
        return candidate
    return "npx.cmd"


class MCPClient:
    """Minimal JSON-RPC stdio client for the AWS Pricing Calculator MCP server."""

    def __init__(self):
        self._proc = None
        self._lock = threading.Lock()
        self._id   = 0

    def _ensure_proc(self):
        if self._proc is None or self._proc.poll() is not None:
            self._proc = subprocess.Popen(
                [_find_npx(), "-y", "sample-aws-pricing-calculator-mcp@latest"],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, bufsize=1, shell=False,
            )
            time.sleep(5)
            self._handshake()

    def _handshake(self):
        self._rpc("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "sow-agent", "version": "1.0"},
        })
        req = {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}}
        self._proc.stdin.write(json.dumps(req) + "\n")
        self._proc.stdin.flush()

    def _rpc(self, method: str, params: dict) -> dict:
        self._id += 1
        req = {"jsonrpc": "2.0", "id": self._id, "method": method, "params": params}
        self._proc.stdin.write(json.dumps(req) + "\n")
        self._proc.stdin.flush()
        for _ in range(600):
            line = self._proc.stdout.readline().strip()
            if not line:
                time.sleep(0.05)
                continue
            try:
                resp = json.loads(line)
                if resp.get("id") == self._id:
                    if "error" in resp:
                        raise RuntimeError(str(resp["error"]))
                    return resp.get("result", {})
            except json.JSONDecodeError:
                continue
        raise TimeoutError(f"No response for {method}")

    def call_tool(self, tool_name: str, arguments: dict) -> str:
        """Call an MCP tool and return its text content."""
        with self._lock:
            self._ensure_proc()
            result  = self._rpc("tools/call", {"name": tool_name, "arguments": arguments})
            content = result.get("content", [])
            return "\n".join(c.get("text", "") for c in content if c.get("type") == "text")

    def list_tools(self) -> list[dict]:
        """Return available tool definitions from the MCP server."""
        with self._lock:
            self._ensure_proc()
            result = self._rpc("tools/list", {})
            return result.get("tools", [])


_mcp = MCPClient()


# ─── Bedrock agent loop ───────────────────────────────────────────────────────

def _get_bedrock_client():
    return boto3.client(
        "bedrock-runtime",
        region_name=os.getenv("AWS_REGION", "ap-south-1"),
        aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
        aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
    )


def _build_tool_spec() -> list[dict]:
    """
    Fetch tool definitions from the MCP server and convert to
    Bedrock converse API toolSpec format.
    """
    mcp_tools = _mcp.list_tools()
    specs = []
    for t in mcp_tools:
        # Build inputSchema — use MCP schema if present, else accept any object
        schema = t.get("inputSchema") or {"type": "object", "properties": {}}
        specs.append({
            "toolSpec": {
                "name":        t["name"],
                "description": t.get("description", ""),
                "inputSchema": {"json": schema},
            }
        })
    return specs


def _run_agent(system_prompt: str, user_message: str, max_turns: int = 20) -> str:
    """
    Agentic loop using Bedrock Converse API with tool use.
    The model decides which MCP tools to call and in what order.
    Returns the final text response from the model.
    """
    bedrock    = _get_bedrock_client()
    tool_specs = _build_tool_spec()
    messages   = [{"role": "user", "content": [{"text": user_message}]}]

    print(f"[Agent] Starting with {len(tool_specs)} tools available")

    for turn in range(max_turns):
        response = bedrock.converse(
            modelId="global.anthropic.claude-haiku-4-5-20251001-v1:0",
            system=[{"text": system_prompt}],
            messages=messages,
            toolConfig={"tools": tool_specs},
        )

        output      = response["output"]["message"]
        stop_reason = response["stopReason"]
        messages.append(output)

        print(f"[Agent] Turn {turn+1} stopReason={stop_reason}")

        if stop_reason == "end_turn":
            # Extract final text
            for block in output.get("content", []):
                if "text" in block:
                    return block["text"]
            return ""

        if stop_reason == "tool_use":
            tool_results = []
            for block in output.get("content", []):
                if "toolUse" not in block:
                    continue
                tool_call = block["toolUse"]
                name      = tool_call["name"]
                tool_id   = tool_call["toolUseId"]
                args      = tool_call.get("input", {})

                print(f"[Agent] → {name}({list(args.keys())})")
                try:
                    result_text = _mcp.call_tool(name, args)
                    print(f"[Agent]   ✓ {result_text[:120]}")
                    tool_results.append({
                        "toolResult": {
                            "toolUseId": tool_id,
                            "content":   [{"text": result_text}],
                            "status":    "success",
                        }
                    })
                except Exception as e:
                    err = str(e)
                    print(f"[Agent]   ✗ {err[:120]}")
                    tool_results.append({
                        "toolResult": {
                            "toolUseId": tool_id,
                            "content":   [{"text": f"Error: {err}"}],
                            "status":    "error",
                        }
                    })

            messages.append({"role": "user", "content": tool_results})

    raise TimeoutError(f"Agent did not finish in {max_turns} turns")


# ─── Service table extraction (Bedrock, no tools needed) ─────────────────────

def _detect_region(text: str) -> str:
    t = text.lower()
    if "us-east" in t or "virginia" in t: return "us-east-1"
    if "us-west" in t or "oregon" in t:   return "us-west-2"
    if "ireland" in t or "eu-west" in t:  return "eu-west-1"
    if "singapore" in t:                  return "ap-southeast-1"
    if "sydney" in t:                     return "ap-southeast-2"
    if "tokyo" in t:                      return "ap-northeast-1"
    return "ap-south-1"


def _parse_pipe_table(raw: str) -> list[dict]:
    services = []
    for line in raw.strip().split("\n"):
        line = line.strip()
        if not line or "|" not in line or line.startswith("```"):
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 4:
            continue
        name = parts[0].strip("*- ").strip()
        if not name or name.lower() in ("service_name", "service", "name"):
            continue
        services.append({
            "name": name,
            "dev":  {"spec": parts[1].strip() or "—"},
            "uat":  {"spec": parts[2].strip() or "—"},
            "prod": {"spec": parts[3].strip() or "—"},
        })
    return services


def extract_services_from_mom(customer_name: str, mom_text: str) -> dict:
    """
    Ask Bedrock to read the MOM and return a pipe-delimited service table.
    Returns { region, services: [{name, dev:{spec}, uat:{spec}, prod:{spec}}] }
    """
    from bedrock_client import call_bedrock

    prompt = f"""You are an AWS Solutions Architect. Read this Meeting of Minutes and identify every AWS service needed.

Customer: {customer_name}
MOM:
{mom_text}

Output ONLY a pipe-delimited table. One AWS service per line. Exactly 4 columns. NO header row.
Format: SERVICE_NAME | DEV_SPEC | UAT_SPEC | PROD_SPEC Prod en

- SERVICE_NAME: exact AWS service name
- DEV_SPEC / UAT_SPEC / PROD_SPEC: short sizing description based on MOM (use — if not needed)
- No markdown. No explanation. Just the pipe-delimited lines."""

    raw      = call_bedrock(prompt)
    services = _parse_pipe_table(raw)

    if not services:
        # Retry simpler
        raw2     = call_bedrock(f"List AWS services for: {mom_text[:500]}\nOne per line: SERVICE | DEV | UAT | PROD")
        services = _parse_pipe_table(raw2)

    if not services:
        raise ValueError("Bedrock could not extract services from the MOM. Please add more detail.")

    print(f"[Bedrock] {len(services)} services extracted from MOM")
    return {"region": _detect_region(mom_text), "services": services}


# ─── Main entry point ─────────────────────────────────────────────────────────

def build_pricing_estimate(customer_name: str, mom_text: str, highlights: str = "") -> dict:
    """
    Agent-based pricing estimation:

    1. Bedrock (no tools) reads MOM → service table for SOW display
    2. Bedrock agent loop with MCP tools:
       - Agent decides to call search_services, get_service_fields,
         create_estimate, add_service, validate_estimate, export_estimate
       - Agent handles errors autonomously (field grounding, lint failures)
       - Agent returns the calculator.aws URL
    """
    # Step 1: Extract service table for SOW display
    url          = "https://calculator.aws/pricing/2/estimate"
    services     = []
    region       = "ap-south-1"
    cost_summary = ""
    error        = None

    try:
        extracted = extract_services_from_mom(customer_name, mom_text)
        services  = extracted["services"]
        region    = extracted["region"]
    except ValueError as ve:
        error = str(ve)
        return {
            "url": url, "services": [], "added_services": [],
            "failed_services": [], "cost_summary": "", "region": region, "error": error,
        }

    # Step 2: Agent loop — let Bedrock autonomously build the estimate
    system_prompt = """You are an AWS cost estimation agent. Your job is to build an AWS Pricing Calculator estimate and return a shareable URL.

You have access to these MCP tools:
- search_services: find the correct service key for a service name
- get_service_fields: get field IDs, types, valid options, and minimalConfig for a service
- create_estimate: create a new empty estimate, returns estimate_id
- add_service: add a service to an estimate using correct field IDs
- validate_estimate: check if estimate is valid before exporting
- export_estimate: save and get the shareable calculator.aws URL

Rules:
- Always call get_service_fields before add_service to discover the correct field IDs and use minimalConfig as your starting point
- If add_service returns field errors, call get_service_fields again and fix the fields
- If validate_estimate returns lint errors, fix the affected services before exporting
- Only call export_estimate when validate_estimate returns lint_verdict: editable
- End your response with the final calculator.aws URL on its own line"""

    service_list = "\n".join(
        f"- {s['name']}: {s['prod'].get('spec', '')} (region: {region})"
        for s in services
        if s.get("prod", {}).get("spec", "—") != "—"
    )

    user_message = f"""Build an AWS Pricing Calculator estimate for {customer_name}.

Region: {region}
Estimate name: SOW - {customer_name}

Services to include (production configuration):
{service_list}

Instructions:
1. For each service, call search_services to find its key
2. Call get_service_fields to get the minimalConfig
3. Create the estimate with create_estimate
4. Add each service with add_service using minimalConfig fields
5. Call validate_estimate — if editable, call export_estimate
6. Return the shareable URL"""

    try:
        print(f"\n[Agent] Starting pricing agent for {len(services)} services")
        agent_response = _run_agent(system_prompt, user_message)
        print(f"[Agent] Final response: {agent_response[:300]}")

        # Extract URL from agent response
        m = re.search(r'https://calculator\.aws[^\s\)\"\]]+', agent_response)
        if m:
            url = m.group(0).rstrip(".,)")
            print(f"[Agent] URL extracted: {url}")

            # Fetch cost summary
            aws_id_m = re.search(r'estimate\?id=([a-f0-9]+)', url)
            if aws_id_m:
                try:
                    cost_summary = _mcp.call_tool("import_estimate", {
                        "estimate_id": aws_id_m.group(1),
                        "format":      "markdown",
                    })[:2000]
                    print(f"[Agent] Cost summary: {len(cost_summary)} chars")
                except Exception as ie:
                    print(f"[Agent] import warning: {ie}")
        else:
            error = "Agent completed but no calculator URL found in response"
            print(f"[Agent] {error}")

    except Exception as e:
        error = str(e)
        print(f"[Agent] Error: {e}")

    return {
        "url":             url,
        "services":        services,
        "added_services":  [],
        "failed_services": [],
        "cost_summary":    cost_summary,
        "region":          region,
        "error":           error,
    }
