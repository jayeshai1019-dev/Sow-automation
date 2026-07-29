import boto3
import json
import re
import os
from dotenv import load_dotenv

# Load .env from the backend folder
load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

def _get_client():
    """Create a Bedrock runtime client using credentials from .env"""
    return boto3.client(
        service_name="bedrock-runtime",
        region_name=os.getenv("AWS_REGION", "ap-south-1"),
        aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
        aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
    )

def call_bedrock(prompt: str) -> str:
    """Call AWS Bedrock Claude Haiku to generate dynamic content."""
    try:
        client = _get_client()
        body = json.dumps({
            "anthropic_version": "bedrock-2023-05-31",
            "max_tokens": 1024,
            "messages": [{"role": "user", "content": prompt}]
        })
        response = client.invoke_model(
            modelId="global.anthropic.claude-haiku-4-5-20251001-v1:0",
            body=body,
            contentType="application/json",
            accept="application/json"
        )
        result = json.loads(response["body"].read())
        return result["content"][0]["text"].strip()
    except Exception as e:
        print(f"Bedrock error: {e}")
        return f"[Content to be filled - Bedrock error: {str(e)}]"


def generate_about_customer(customer_name: str, mom_text: str, company_about_us: str = "") -> str:
    """
    Generate 'About Customer' section (1.3) in 5-6 lines.
    If company_about_us is provided (scraped from URL), use it as primary source.
    Otherwise generate from MOM notes.
    """
    if company_about_us and company_about_us.strip():
        prompt = f"""Rewrite the following company information into a professional 'About Customer' paragraph of exactly 5 to 6 sentences for a Statement of Work document.

Customer Name: {customer_name}
Company Information: {company_about_us}

Rules:
- Write exactly 5 to 6 sentences.
- Professional tone, suitable for a formal SOW document.
- Mention the company name, what they do, their industry or domain, and their scale or key strengths.
- Do NOT include headings, bullet points, or labels.
- Output only the paragraph text."""
    else:
        prompt = f"""Write a professional 'About Customer' paragraph of exactly 5 to 6 sentences for a Statement of Work document, based on the meeting notes below.

Customer Name: {customer_name}
Meeting Notes: {mom_text}

Rules:
- Write exactly 5 to 6 sentences.
- Professional tone, suitable for a formal SOW document.
- Mention the company name, what they do, their industry or domain, and any relevant context from the notes.
- Do NOT include headings, bullet points, or labels.
- Output only the paragraph text."""
    return call_bedrock(prompt)


def generate_project_objectives(customer_name: str, mom_text: str, project_type: str = "POC") -> str:
    engagement_hint = (
        "This is a Proof of Concept (POC) engagement."
        if project_type == "POC"
        else "This is a full Production / standard project engagement — do NOT use POC language."
    )
    prompt = f"""You are drafting the "1.1 Project Objectives and Overview" section of a Statement of Work (SOW) for Operisoft Technologies (the solution provider). Operisoft always refers to itself as "Operisoft" and to the client as "the customer."

Engagement Context: {engagement_hint}

STRICT FORMAT — follow this exact structure, no deviations:
- Exactly 2 bullet points, each starting with "Ø  " (Ø followed by two spaces).
- Bullet 1: A thank-you line in this pattern:
  "Ø  Operisoft would like to thank the team of <Customer Name> (hereby referred to as the customer) for engaging with us on/for <a short, specific description of the engagement type and solution — e.g. Proof of Concept (POC) for a GenAI-based X Solution, or the migration of their existing workloads to AWS cloud>."
- Bullet 2: A description line starting with "Ø  As part of this POC/project, ..." that explains:
  - what is being built or done
  - what inputs/data sources are involved (systems, documents, databases, files mentioned in the MOM)
  - what manual process is being automated or what business outcome is targeted

RULES:
- If the engagement is a POC, use "Proof of Concept (POC)" language in bullet 1 and "As part of this POC" in bullet 2.
- If the engagement is Production, use appropriate project language (e.g., "project", "engagement", "implementation") — do NOT use the word "POC" anywhere.
- Infer the correct engagement type details (migration, automation, chatbot, etc.) from the MOM — but the POC/Production flag above overrides any ambiguity.
- Be specific: name the actual technologies, data sources, and outputs mentioned in the MOM (e.g. S3, MySQL, Excel, PDF, LLM) rather than generic terms.
- No headings, no extra bullets, no closing remarks, no markdown bold.
- Match the formal, concise register of the examples below — do not pad with generic filler sentences.

EXAMPLES OF DESIRED OUTPUT:

Example 1:
Ø  Operisoft would like to thank the team of Aptech (hereby referred to as the customer) for engaging with us on a Proof of Concept (POC) for a GenAI-based PDF Data Extraction and Automation Solution.
Ø  As part of this POC, the Customer intends to automatically extract specific data fields from PDF documents and generate a structured Excel output with required transformations, in order to streamline reporting and eliminate manual processing efforts.

Example 2:
Ø  Operisoft would like to thank the team of Landmark Insurance Brokers Pvt. Ltd. (hereby referred to as the customer) for engaging with us on a Proof of Concept (POC) for a Gen AI - based Employee Support Chatbot Solution.
Ø  As part of this POC, a Gen AI-powered chatbot will be developed to enable employees to raise general queries. The chatbot will leverage the customer's existing MySQL database as its knowledge source, along with customer-provided unique identifiers (via an Excel sheet) to understand and contextualize employee queries.

Example 3 (Production):
Ø  Operisoft would like to thank the team of Acme Corp (hereby referred to as the customer) for engaging with us on the migration of their existing workloads to AWS cloud.
Ø  As part of this project, the customer intends to migrate their on-premises infrastructure — including application servers and databases — to AWS, leveraging services such as Amazon EC2, Amazon RDS, and AWS MGN for a seamless, secure transition.

NOW GENERATE FOR:
Customer Name: {customer_name}
MOM Notes:
{mom_text}

Output only the 2 bullet points, nothing else."""
    return call_bedrock(prompt)


def generate_current_landscape(customer_name: str, mom_text: str) -> str:
    prompt = f"""You are a solutions consultant writing the "Understanding of Current Landscape" section of a formal AWS Statement of Work (SOW) for {customer_name}, based on the meeting notes (MOM) below.

MOM Notes:
{mom_text}

## Critical Rule — No Hallucination
Only state something as "current" or "existing" if the MOM explicitly says it already exists today. 
MOM content that describes what the customer WANTS, NEEDS, or is ASKING FOR (e.g. answers to a requirements questionnaire — server specs, database choice, expected user load, backup frequency, HA/DR preference) is a REQUIREMENT, not a current state. Never rewrite a requirement as if it already exists or is already running somewhere.
If the MOM contains no description of an existing system, application, or infrastructure, explicitly state that no current AWS/cloud environment or application exists — do not invent one.

## Structure to follow (three parts, in this order)

**1. Current Scenario**
- If the MOM describes an existing system/process/infra: describe it factually — what exists, how it works today, tools/vendors in use, locations, scale (numbers), and pain points — but only using facts stated in the MOM.
- If the MOM is a requirements questionnaire for a NEW application/system with nothing currently in place: state plainly that the customer does not currently have this hosted/deployed (e.g., "[Customer] does not currently have an AWS-hosted environment for [application type]; this is a new deployment.") Do not fabricate an on-prem or legacy backstory.

**2. Requirement**
- List what the customer wants/needs, exactly as captured in the MOM — application type, tech stack, database, sizing, user load, access model, backup expectations, HA/DR/security preferences, etc.
- Present as clear sub-bullets. Preserve all specific numbers, names, and technical terms exactly as given.
- Frame these as requirements/asks, not as existing facts (e.g. "The customer requires..." / "Expected concurrent user load: 10,000 users" — not "The application currently supports 10,000 users").

**3. Proposed Approach (high-level only)**
- One or two sentences, high level only — name the type of solution (e.g., new AWS hosting deployment, migration, POC, chatbot) and its overall intent.
- Do NOT include architecture details, specific AWS services, instance types, or design decisions — that belongs in a later section of the SOW.
- Use transition phrasing such as: "To meet these requirements, Operisoft proposes to..."

## Formatting rules
- Use "Ø" for top-level bullets and "·" for sub-bullets
- No markdown headers (no ##, no bold) — output should read as plain SOW bullet text, but you may use the labels "Current Scenario", "Requirement", and "Proposed Approach" as plain-text section labels (not markdown headers)
- Formal, factual, third-person tone — client-facing document
- Do not omit any specific number, name, or technical detail present in the MOM
- If a detail is not present in the MOM, omit it — do not guess or fill gaps

## Example (greenfield / new deployment MOM)

Current Scenario
Ø [Customer] does not currently have an AWS-hosted environment for its [application type] application; this is a new cloud deployment.

Requirement
Ø The customer intends to host a [application type] application on AWS with the following specifications:
· Frontend: [X] | Backend: [Y]
· Database: [type/version]
· Server sizing — Application: [vCPU/RAM], Database: [storage]
· Expected concurrent user load: [N] users
· Access model: [public-facing/internal]
· Backup frequency: [X]
· High availability required: [Yes/No]
· Disaster recovery required: [Yes/No]
· Firewall/WAF required: [Yes/No]

Proposed Approach
Ø To meet these requirements, Operisoft proposes to design and deploy a new, secure AWS hosting environment tailored to the above specifications.

## Example (existing/manual process MOM)

Current Scenario
Ø Currently, there is no automated mechanism to extract and structure required data from the PDF documents.
Ø All data extraction and validation activities are performed manually, including identifying required fields, interpreting remarks, and consolidating data into Excel.
Ø The current manual process is time-consuming, repetitive, and prone to human error.

Requirement
Ø The customer requires an automated solution to extract, validate, and structure data from PDF documents stored in Amazon S3, applying defined business rules and producing a structured Excel output.

Proposed Approach
Ø To address these challenges, Operisoft proposes a Proof of Concept (POC) for a GenAI-based PDF data extraction solution using a Large Language Model (LLM).

## Output
Generate ONLY the three labeled sections (Current Scenario / Requirement / Proposed Approach) with bulleted content as shown above. No preamble, no extra commentary, no heading beyond the three plain-text labels.
"""
    return call_bedrock(prompt)


def generate_key_highlights(customer_name: str, mom_text: str) -> str:
    prompt = f"""Based on the following MOM notes, generate a bullet-point list of 'Key Highlights of the Solution' for an AWS Cloud Migration SOW.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Return only bullet points (one per line starting with •), covering AWS services, architecture choices, regions, and migration approach. Keep each point concise."""
    return call_bedrock(prompt)


def generate_dr_requirements(customer_name: str, mom_text: str) -> dict:
    prompt = f"""Based on the following MOM notes, fill in the DR (Disaster Recovery) requirements for an AWS SOW.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Return a JSON object with these exact keys:
- critical_systems: description of critical systems
- primary_site: primary AWS region
- dr_type: type of DR (Active/Passive/Warm Standby)
- rto: Recovery Time Objective (e.g. 30 minutes)
- rpo: Recovery Point Objective (e.g. 15 minutes)

Return ONLY the JSON object, no other text."""
    result = call_bedrock(prompt)
    try:
        json_match = re.search(r'\{.*\}', result, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
    except Exception:
        pass
    return {
        "critical_systems": "All production servers and databases",
        "primary_site": "Mumbai Region (ap-south-1)",
        "dr_type": "Passive",
        "rto": "30 minutes",
        "rpo": "15 minutes"
    }


def generate_proposed_drs_solution(customer_name: str, mom_text: str) -> str:
    prompt = f"""Based on the following MOM notes, write the 'Proposed DRS Solution' section for an AWS SOW describing the disaster recovery strategy.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Write 4-6 bullet points describing the DR solution. Start each bullet with •. Cover replication approach, networking, RTO/RPO targets, and failover process."""
    return call_bedrock(prompt)


def generate_tco_assumptions(customer_name: str, mom_text: str, cost_summary: str = "") -> str:
    services_context = ""
    if cost_summary and cost_summary.strip():
        services_context = f"\n\nActual AWS services included in the pricing estimate (use these as the basis for assumptions):\n{cost_summary[:1500]}"

    prompt = f"""Based on the following MOM notes, write the 'Assumptions Taken for Cost Calculation' section for an AWS SOW.
Customer Name: {customer_name}
MOM Notes: {mom_text}{services_context}

Write 6-8 bullet points (starting with •) about cost assumptions.
{"Since the actual AWS services in the estimate are provided above, make the assumptions specific to those services (e.g. EC2 instance types, S3 storage tiers, RDS engine, Lambda invocation count, regions, backup retention)." if cost_summary else "Cover regions, compute sizing, backup, storage, and uptime."}
Each assumption should be concrete and measurable — avoid vague generic statements.
Be specific where information is available from the MOM or the service list."""
    return call_bedrock(prompt)


def generate_high_level_scope(customer_name: str, mom_text: str) -> str:
    prompt = f"""You are a solutions consultant writing the "High Level Scope" section (3.1) of a formal AWS Statement of Work (SOW) for {customer_name}, based on the meeting notes (MOM) below.

MOM Notes:
{mom_text}

## Task
Describe, at a high level, the solution Operisoft will implement and its key capabilities — grounded strictly in the MOM. This is a high-level scope statement, not a detailed architecture document.

## Structure to follow
1. Opening line: "Ø Operisoft will implement a/an [Solution Type Name], capable of:" 
   - Name the solution type crisply and specifically (e.g., "GenAI-based PDF Data Extraction and Automation Solution", "AI-powered NoSQL Chatbot Solution") — based on what the MOM describes the customer wants.
2. List the solution's capabilities as sub-bullets, in the natural order data/users would flow through the system:
   - Entry point first (UI, upload mechanism, chatbot interface, authentication)
   - Then storage (where data lands)
   - Then processing/orchestration (what triggers and coordinates the workflow)
   - Then the AI/LLM layer (what it analyzes/interprets and how)
   - Then output/delivery (dashboard, export, download, response to user)
   If there are more distinct components than fit within the limit below, merge closely related ones into a single bullet (e.g., combine storage + processing into one line) rather than dropping a component entirely.
3. Each bullet must name the specific AWS service responsible for that capability wherever the MOM implies or states one. If the MOM doesn't specify a service for a needed capability, choose the standard/appropriate AWS service for that function — do not invent obscure or unlikely services.
4. Optionally, end with a "Note:" line (plain text, not a bullet) if there's an important scope boundary to clarify. Only include this if relevant; do not force it. The Note line also counts toward staying concise — max 2 lines.

## STRICT LENGTH CONSTRAINTS
- Maximum 5 capability bullets total. Never exceed 5. If there are more than 5 real components, consolidate the least critical or most closely related ones together.
- Each bullet must be a maximum of 2 lines when rendered (roughly 25-30 words). Do not write long, multi-clause sentences — keep each bullet tight and single-purpose. If a capability needs more explanation than that, trim it to its essential point.

## Formatting rules
- Use "Ø" for the opening line and "·" for each capability sub-bullet, with consistent spacing (two spaces after the symbol)
- No markdown headers, no bold/italics
- Formal, third-person, concise
- Do NOT include implementation details like specific instance sizes, IAM policies, network CIDR ranges, or step-by-step logic
- Do NOT invent capabilities, services, or features not implied by the MOM

## Examples of expected style, depth, and length

Example A (document extraction):
Ø Operisoft will implement a GenAI-based PDF Data Extraction and Automation Solution, capable of:
·  A web-based UI for uploading PDF documents, hosted on AWS Amplify.
·  Secure storage of uploaded PDFs in Amazon S3, processed via AWS Lambda to orchestrate the extraction workflow.
·  AI-driven data extraction using a Large Language Model via Amazon Bedrock to identify fields and apply business rules.
·  Generation of a structured Excel output, downloadable through the front-end interface.

Example B (chatbot solution, with scope-boundary Note):
Ø Operisoft will implement an AI-powered NoSQL Chatbot Solution as part of the project, capable of:
·  Allowing users to ask questions through a sample chatbot UI using natural language.
·  Maintaining chat history for the session to support conversational continuity.
·  Interpreting user questions using AI-based context understanding via Amazon Bedrock.
·  Presenting clear, human-readable responses via the chatbot interface.
Note: The sample UI includes only basic functionality to demonstrate conversational data access.

## Output
Generate ONLY the "High Level Scope" section content in the above style — the opening Ø line, up to 5 · sub-bullets (each max 2 lines), and an optional Note line. No heading, no preamble, no explanation.
"""
    return call_bedrock(prompt)


def generate_scope_of_work(customer_name: str, mom_text: str, project_type: str = "POC") -> str:
    engagement_hint = (
        "POC (Proof of Concept)"
        if project_type == "POC"
        else "Production / full-scale project"
    )
    prompt = f"""You are a solutions consultant writing the "Scope of Work" section (3.2) of a formal AWS Statement of Work (SOW) for {customer_name}, based on the meeting notes (MOM) below.

Engagement Type (explicit, from user): {engagement_hint}

MOM Notes:
{mom_text}

## Step 1 — Determine engagement type FIRST
The engagement type has been explicitly provided above as "{engagement_hint}". Use this to determine POC vs Production scope.
- If it is "POC (Proof of Concept)": use POC language throughout (e.g. "As part of the POC, Operisoft will..."), and include POC-specific disclaimer notes.
- If it is "Production / full-scale project": use standard project language (e.g. "As part of this project, Operisoft will..."), do NOT use the word "POC" anywhere, and omit POC-specific disclaimers.
- This explicit flag overrides any wording in the MOM — even if the MOM mentions "POC", follow the explicit flag.

## Step 2 — Write the Scope of Work
Describe, in narrative bullet form, what Operisoft will build and how the solution will work end-to-end, following the natural flow of the system:
1. What the customer will provide/share (sample data, access, credentials, existing systems) — if mentioned in the MOM
2. How users will access the solution — UI, authentication method, hosting (e.g., AWS Amplify, Cognito) — if applicable
3. How requests/data are routed and orchestrated (API Gateway, Lambda functions) and what each major processing step does
4. How data is stored at each stage (landing bucket, processing bucket, response bucket, database) — if the MOM implies multiple storage stages, describe each
5. AI/LLM-driven processing — what it analyzes and what business logic/rules it applies. If there are multiple distinct validation/extraction steps, list them as sub-bullets under this point (only do this if the MOM's process genuinely has multiple distinct steps — don't force sub-bullets for a simple flow)
6. Final output and how the user receives/views it (dashboard, Excel export, download, chatbot response)
7. A dedicated "Services & Integration" bullet listing all specific AWS services used, e.g.: "Ø Services & Integration: · The solution will utilize [service list]."
8. A closing bullet on collaborative validation/testing between Operisoft and the customer team

Match the DEPTH and COMPLEXITY of the bullets to what the MOM actually describes:
- If the MOM describes a simple, flat process (e.g., a basic chatbot), keep bullets flat and concise — don't invent extra steps.
- If the MOM describes a complex multi-stage pipeline (e.g., multiple Lambda functions, multiple validation checks), reflect that complexity using nested sub-bullets where appropriate.
- The number of top-level bullets should reflect the real number of distinct stages in the MOM — do not force a fixed count.

## Step 3 — Write the Notes section
Add a "Notes:" section (plain text label, not a bullet) with bullets covering:
- IF this is a POC: include the accuracy disclaimer ("Operisoft will build a solution based on customer's requirements, but 100% accuracy of the tool is not guaranteed"), the assumptions/future-services disclaimer, and the production-SOW disclaimer ("For production-level implementation or a POC with additional requirements, a new SOW will be defined with applicable additional charges").
- IF this is Production: only include the assumptions/future-services disclaimer ("This solution is based on certain assumptions, and additional AWS services may need to be incorporated in the future for this requirement."). Omit the POC accuracy disclaimer and the "new SOW for production" line.

## Formatting rules
- Use "Ø  " (with two spaces) for top-level bullets and "·        " for sub-bullets, matching standard SOW bullet spacing
- No markdown headers, no bold/italics
- Formal, third-person, technical but readable tone
- Name specific AWS services exactly as implied/stated in the MOM (e.g., Amazon S3, AWS Lambda, Amazon Bedrock, AWS Amplify, Amazon Cognito, Amazon API Gateway, Amazon DynamoDB, AWS CloudFormation) — do not invent unlikely services
- Do NOT include pricing, timelines, or team/resourcing details — those belong in other SOW sections
- Do NOT fabricate steps, services, or capabilities not implied by the MOM

## Examples of expected style, depth, and structure

Example A (POC, simple flat flow — chatbot):
Ø  For the project, the Customer will share sample key–value data, provide read-only access to the NoSQL database, and grant appropriate access to the Customer's AWS account to enable deployment and configuration of the required services during the project.
Ø  A sample chatbot UI will be developed exclusively for this. The UI will include only the basic functionality required to ask questions and view responses and will not represent a production-grade application.
Ø  The chatbot will retrieve relevant information needed to respond to the question from the Customer's NoSQL data in a read-only manner, using Generative AI to understand intent and generate natural language responses.
Ø  Services & Integration: Operisoft will leverage and deploy AWS services including Amazon API Gateway, AWS Lambda, Amazon S3, Amazon Bedrock, and Amazon EC2, provisioned and managed through AWS CloudFormation.
Ø  Operisoft and the Customer team will collaboratively perform validation and functional testing to confirm that the project meets the agreed objectives.

Notes:
Ø  Operisoft will build a solution based on customer's requirements, but 100% accuracy of the tool is not guaranteed.
Ø  This solution is based on certain assumptions, and additional AWS services may need to be incorporated in the future for this requirement.

Example B (POC, complex multi-stage flow — document verification, with nested sub-bullets):
Ø  Authorized users will access a web-based User Interface (UI) hosted on AWS Amplify, with authentication and access control managed through Amazon Cognito.
Ø  Users will upload documents through the UI. All requests will be routed via Amazon API Gateway to backend services, with AWS Lambda handling request initialization, metadata capture, and orchestration. Documents will be securely stored in Amazon S3, with processing status tracked using Amazon DynamoDB.
Ø  A Validation Lambda will invoke the LLM (via Amazon Bedrock) to perform:
·        Document completeness validation
·        Verification of correct document types
·        Extraction of relevant data fields
·        Cross-verification of form data against uploaded documents
Ø  The processed and verified results will be stored in an Amazon S3 Response Bucket and made available on a Dashboard hosted on AWS Amplify, with results downloadable in Excel format.
Ø  Services & Integration: The solution will utilize Amazon S3, AWS Lambda, Amazon API Gateway, Amazon Bedrock, AWS Amplify, Amazon Cognito, and Amazon DynamoDB.
Ø  Operisoft and the customer team will collaboratively perform validation and functional testing to ensure the solution meets the defined POC objectives.

Notes:
Ø  As this is a POC project, Operisoft will build a solution based on customer's requirements, but 100% accuracy of the tool is not guaranteed.
Ø  This solution is based on certain assumptions, and additional AWS services may need to be incorporated in the future for this requirement.
Ø  For production-level implementation or a POC with additional requirements, a new SOW will be defined with applicable additional charges.

## Output
Generate ONLY the "Scope of Work" bulleted content followed by the "Notes:" section, exactly as shown in structure above. No section heading for "Scope of Work" itself, no preamble, no explanation.
"""
    return call_bedrock(prompt)


def generate_customer_dependencies(customer_name: str, mom_text: str) -> str:
    """Generate 'Customer Dependencies' section from MOM notes."""
    prompt = f"""You are a solutions consultant writing the "Customer Dependencies" section of a formal AWS Statement of Work (SOW) for {customer_name}.

MOM Notes:
{mom_text}

## Important Context
This section draws from a STANDARD POOL of dependency clauses used across Operisoft SOWs. Your job is to select and reproduce the clauses from this pool that are relevant to THIS project, based on the MOM — not to freely generate a custom list from scratch.

## Step 1 — Determine if this is a migration / on-prem-to-AWS project
Check the MOM for signals of data migration, existing on-premises infrastructure being connected to or migrated to AWS, or network connectivity between a customer data centre and AWS (e.g., mentions of source data centre, existing servers, VPN, bandwidth, agents, network architecture).
- If YES (migration/on-prem involved): include the full pool below, including the network/bandwidth/migration-specific clauses.
- If NO (e.g., a new greenfield application deployment, a POC, a chatbot, a document-processing solution with no existing on-prem system): OMIT the network architecture, bandwidth, and network-alignment clauses — they will not make sense for this type of project. Include only the clauses that are universally applicable (stakeholder availability, application configuration responsibility, licenses, code/config changes, access control).

## Standard Clause Pool

CRITICAL — Reproduce these EXACTLY as written below, character-for-character. Do NOT paraphrase, rewrite, expand, or rephrase them, even if alternate wording seems better. Do not change "will" to "shall" or vice versa. Do not add extra sentences or qualifying phrases to any core clause. These must read identically across all SOWs.

Always applicable (include in every SOW, unless clearly irrelevant to a specific point):
Ø  Availability of the key/right stakeholders will be assured for defined activities MRA / Discovery Session.
Ø  All application-level configuration changes to be done by the customer in-house team and AWS PARTNER will provide support related to AWS-specific configurations.
Ø  Licenses for all applications and third-party tools will be provided by the customer.
Ø  Any changes required in application code/configuration or database configuration will be done by the customer's team.
Ø  Providing access control details for IAM, S3, Security Groups as per business requirement.

Only include if this is a migration / on-prem-to-AWS project (per Step 1):
Ø  The customer team will provide the details of network architecture and application dependencies.
Ø  The customer will have to ensure the availability of sufficient network bandwidth at the source data centre for the success of the project as per the defined timelines in this document. It is advised to start the project only when the required bandwidth is provisioned.
Ø  Bandwidth requirements will be mutually agreed upon during the initial phase of the feasibility of migration methodology for the database and other components.
Ø  All third-party tools and services will be migrated as-is from the current data centre. If any configuration changes are required, those will be done by the Customer team and AWS PARTNER will provide support wherever required on the AWS cloud infrastructure side.
Ø  Installation of agents for application migration service will be done by the customer's IT team.
Ø  The customer team has to share the Applications dependencies along with compatibility with Operating System licenses etc.
Ø  Customer needs to align their IT and network teams for setting up the network architecture between AWS and their on-premises environment.

## Step 2 — Add MOM-specific dependencies (if any)
After selecting from the pool above, check if the MOM reveals a genuinely unique, project-specific dependency not already covered — e.g., providing specific sample data, specific credentials, specific documents named in the MOM, or specific vendor coordination. Add at most 2-3 such bullets only if clearly implied and non-redundant with the pool above. If none apply, add nothing extra.

## Rules
- Total output should typically be 5-12 bullets depending on project type (fewer for non-migration projects, more for migration projects) — do not force a fixed count.
- Do NOT invent generic dependencies beyond the pool and genuine MOM-specific items.
- Use formal, third-person tone suitable for a contractual document.
- Do NOT include pricing, timelines, or Operisoft's own deliverables.
- Start each bullet with "Ø  " (two spaces).
- Output as a flat bullet list.

## Output
Generate ONLY the bullet points (one per line, starting with "Ø  "). No headings, no preamble, no explanation."""
    return call_bedrock(prompt)


def generate_customer_obligations(customer_name: str, mom_text: str) -> str:
    """Generate 'Customer Obligations and Engagement Terms' section from MOM notes."""
    prompt = f"""You are a solutions consultant writing the "Customer Obligations and Engagement Terms" section of a formal AWS Statement of Work (SOW) for {customer_name}.

MOM Notes:
{mom_text}

## Important Context
This section is largely STANDARD across Operisoft SOWs — it uses a fixed set of core engagement clauses that apply to almost every project, not a fully custom list derived from the MOM. Your job is NOT to generate a long custom list. It is to reproduce the standard core clauses (below), and only ADD a small number of MOM-specific bullets if the MOM reveals a genuinely unique dependency or obligation not already covered by the standard clauses.

## Standard Core Clauses

CRITICAL — Reproduce these EXACTLY as written below, character-for-character. Do NOT paraphrase, rewrite, expand, or rephrase them, even if alternate wording seems better. Do not change "will" to "shall" or vice versa. Do not add extra sentences or qualifying phrases to any core clause. These must read identically across all SOWs.

Ø  Inputs provided from the client in any terms will be assumed to be accurate.
Ø  AWS usage and billing will be borne by client during the development as well as at UAT and deployment.
Ø  Necessary approvals and timely review of submissions.
Ø  If a need of any new plugin etc. or any 3rd party service provider needed – will be arranged by client.
Ø  Customer will complete testing of all the deployed AWS services before handover.
Ø  Operisoft will provide consultation for AWS services design and deployment strategies and execute mutually agreed upon new change order.
Ø  Operisoft will not be responsible for any technology or product limitations hindering the service delivery, support or timelines. For any third-party software including any configuration, maintenance or performance issues, anticipated benefits, limitations, expected functionality, compliance or security, the OEM/third party should be contacted.

## Optional MOM-Specific Additions
After the core clauses, review the MOM for any genuinely unique customer obligation not already covered above — for example:
- Customer must provide specific access (e.g., read-only DB access, sample data, VPN access to on-prem systems)
- Customer must nominate a specific point of contact or team for a specific purpose
- Customer must provide specific documents/files/credentials mentioned by name in the MOM
- Any data residency, compliance, or region-specific customer responsibility explicitly mentioned in the MOM

Only add such a bullet if it's clearly implied by the MOM and NOT already covered by the core clauses above (e.g., don't add a generic "customer will provide accurate inputs" bullet — that's already covered). Add at most 2-3 such bullets. If nothing MOM-specific and non-redundant is found, add none — it is completely fine to output only the core clauses.

## Rules
- Do NOT expand this into a long 15-25 bullet list. Total output should be 6-10 bullets: the 7 core clauses (or fewer, if one is clearly inapplicable) plus 0-3 MOM-specific additions.
- Do NOT invent generic "best practice" obligations, exclusions, or assumptions beyond the core clauses and genuine MOM-specific items — this is a lean, standard section, not a comprehensive legal terms document.
- Use formal, third-person tone suitable for a legal/contractual document.
- Do NOT include pricing, timelines, or Operisoft's deliverables here.
- Start each bullet with "Ø  " (two spaces).
- Output as a flat bullet list, core clauses first, then any MOM-specific additions after.



## Output
Generate ONLY the bullet points (one per line, starting with "Ø  "). No headings, no preamble, no explanation."""
    return call_bedrock(prompt)


def generate_assumptions(customer_name: str, mom_text: str) -> str:
    """Generate 'Assumptions' section from MOM notes."""
    prompt = f"""You are a solutions consultant writing the "Assumptions" section of a formal AWS Statement of Work (SOW) for {customer_name}.

MOM Notes:
{mom_text}

## Important Context
This section draws from a STANDARD POOL of assumption clauses used across Operisoft SOWs. Your job is to select and reproduce the clauses relevant to THIS project, plus add a small number of project-specific technical assumptions from the MOM — not to freely generate a custom list from scratch.

## CRITICAL — Reproduce Core Clauses Verbatim
Copy the clauses below EXACTLY as written, character-for-character. Do NOT paraphrase, rewrite, expand, or rephrase them, even if alternate wording seems better. Do not add extra sentences or qualifying phrases to any clause. These must read identically across all SOWs.

## Core Clauses (always include, unless clearly inapplicable):
Ø  A single point of contact will be assigned from the customer team end for coordination and providing the required information, access, approvals, and sign-off.
Ø  All required information and approvals from the customer team will be provided within 3 working days.
Ø  The customer will provide the complete know-how on the application deployment architecture, access, and data flows to Operisoft team.
Ø  Inputs provided from the client in any terms will be assumed to be accurate.
Ø  Necessary approvals and timely review of submissions.
Ø  If a need of any new plugin etc. or any 3rd party service provider needed – will be arranged by client.
Ø  Customer will complete testing of all the deployed AWS services before handover.
Ø  Operisoft will not be responsible for any technology or product limitations hindering the service delivery, support or timelines. For any third-party software including any configuration, maintenance or performance issues, anticipated benefits, limitations, expected functionality, compliance or security, the OEM/third party should be contacted.
Ø  Operisoft's timely and adequate performance of the services and provision of the deliverables shall depend upon full access to appropriate customer personal and to customer information and documentation.
Ø  All the organizational or technical changes which may affect any services will be communicated in writing to Operisoft at least 1 week prior to changes.
Ø  We assume for business-critical applications necessary HA architecture in place to meet the requirements.

## Step 1 — Determine if this involves existing infrastructure, servers, or migration
Check the MOM for signals of: existing servers/infrastructure being sized against, data migration, existing production systems, or ongoing AWS billing responsibility across dev/UAT/production phases.
- If YES: also include this group of clauses:
Ø  Account structure and network architecture will be finalized with customer's IT teams.
Ø  Compute sizing is designed based on the utilization report and existing server configuration.
Ø  AWS usage and billing will be borne by client during the development as well as at UAT and deployment.
Ø  In case of sudden spikes in volumes, caused due to unforeseen circumstances such as major outages, virus incidents, multiple simultaneous failures etc., the support will be on a best effort basis.
Ø  Unauthorized changes: In the event of unauthorized changes in AWS infrastructure, Operisoft reserves rights to refuse support.
- If NO (e.g., new greenfield application, POC, chatbot, document-processing solution with no existing servers/infra to size against): OMIT this group entirely.

If the MOM specifically mentions an AWS environment assessment or migration readiness activity, also include:
Ø  The customer will provide required access to their environments for assessment for AWS migration.

## Step 2 — Add MOM-specific technical assumptions
Add 2-4 bullets stating specific technical assumptions drawn directly from the MOM — e.g., database engine/version, OS, server sizing, region, backup frequency, concurrent user load, or any other concrete technical parameter mentioned. Phrase these as assumptions, e.g., "The database engine will be MySQL and MongoDB as specified by the customer." Only include what's genuinely stated in the MOM — do not invent technical details.

## Rules
- Do NOT paraphrase the core/conditional clauses — reproduce them exactly.
- Total output should typically be 13-18 bullets: the applicable core + conditional clauses, plus 2-4 MOM-specific technical assumptions.
- Use formal, third-person tone suitable for a contractual document.
- Start each bullet with "Ø  " (two spaces).
- Do NOT include pricing details or Operisoft's own deliverables.

## Output
Generate ONLY the bullet points (one per line, starting with "Ø  "). No headings, no preamble, no explanation."""
    return call_bedrock(prompt)


def generate_exclusions(customer_name: str, mom_text: str, project_type: str = "Production") -> str:
    """Generate 'Exclusions' section from MOM notes."""
    prompt = f"""You are a solutions consultant writing the "Exclusions" section of a formal AWS Statement of Work (SOW) for {customer_name}, Project Type: {project_type}.

MOM Notes:
{mom_text}

## Task
Generate a list of items that are explicitly NOT included in Operisoft's scope of work. Exclusions define the boundary of the engagement — they should name things a reasonable customer might assume are included, but aren't, based on what this specific project actually covers.

## Step 1 — Identify the project type/domain from the MOM
Determine what kind of engagement this is (e.g., application deployment, data migration, POC/GenAI solution, chatbot, infrastructure setup) — this determines which domain-specific exclusions are relevant. Use the reference groups below as a guide, and select/adapt whichever group(s) most closely match this MOM. Do not include exclusions from an irrelevant domain (e.g., don't mention "on-premises environment changes" for a pure greenfield cloud app with no on-prem component).

Reference exclusion patterns by project type:

For application/chatbot/POC solutions:
- Troubleshooting or errors related to third-party tools
- Code changes or configurations beyond the defined workflow
- Handling or processing of unsupported data formats or corrupted records
- Development of a production-grade, enterprise-ready solution beyond the defined POC scope (if POC)
- Processing of large-scale datasets or bulk files outside the sample files provided (if POC)
- Custom enhancements, additional business rules, or workflow modifications outside defined requirements
- Any additional requirements apart from this will be considered a separate project, with a separate SOW created for it

For migration/infrastructure projects:
- Security configurations and settings outside of the AWS environment
- Security configuration of the existing environment or corporate network environment
- Any change in OS, application, or third-party tool/service configuration
- Changes in application code or configuration
- 3rd party app integrations
- Modernization activities (to be done after migration phase, as a separate engagement)
- Any change in the current on-premises environment
- Application testing (customer's responsibility)

For deployment projects:
- Security configurations or OS/application/third-party changes beyond the defined deployment scope
- Application code updates, feature changes, or configuration modifications beyond deployment tasks
- Integration with third-party applications or external systems not defined in scope
- Application-level functional or user acceptance testing (customer's responsibility)
- Advanced security assessments beyond agreed scope (e.g., third-party audits, compliance certifications), unless explicitly included

Adapt the wording of whichever bullets apply to reflect this specific MOM's actual solution (e.g., name the specific document type, application, or system involved) rather than copying generically.

## Step 2 — Always close with this standard trio (reproduce verbatim, do not paraphrase)
Ø  Any AWS services training.
Ø  Anything not explicitly mentioned in the scope will be considered out of scope.
Ø  24/7 Technical Support.

(Omit "24/7 Technical Support" only if the MOM explicitly describes a large infrastructure/migration engagement where support terms are covered elsewhere — otherwise always include it.)

## Rules
- Total output should be 6-10 bullets: domain-relevant exclusions (Step 1) + the standard closing trio (Step 2).
- Do NOT include every reference pattern from every domain — only what fits this specific MOM's project type.
- Do NOT repeat items already covered in the Customer Dependencies or Customer Obligations sections.
- Use formal, third-person tone.
- Start each bullet with "Ø  " (two spaces).
- Begin the output with the line: "The following activities are out of scope for this assignment:"

## Output
Generate the header line, followed by the bullet points (one per line, starting with "Ø  "). No additional headings, no preamble, no explanation."""
    return call_bedrock(prompt)


def generate_acceptance_criteria(customer_name: str, mom_text: str, project_type: str = "Production", scope_of_work_text: str = "") -> str:
    type_label = "POC" if project_type.strip().upper() == "POC" else "Production"
    section_heading = "POC Acceptance Criteria" if type_label == "POC" else "3.5 Project Acceptance Criteria"

    scope_context = f"\n\nScope of Work (already generated for this SOW, use as the basis for criteria):\n{scope_of_work_text}" if scope_of_work_text else ""

    prompt = f"""You are a solutions consultant writing the "{section_heading}" section of a formal AWS Statement of Work (SOW) for {customer_name}, Project Type: {type_label}.

MOM Notes:
{mom_text}{scope_context}

## Task
Write the acceptance criteria that define what must be successfully demonstrated for this project/POC to be considered complete and accepted by the customer.

## Critical Rule — Derive from Scope, Don't Invent
Each acceptance criterion should correspond to something the solution is actually being built to do — based on the MOM{" and the Scope of Work provided above" if scope_of_work_text else ""}. Do not use a fixed generic checklist (e.g., don't include "backup and sync" criteria for a chatbot project, or "PDF extraction" criteria for a networking/infra project). Convert each real capability/deliverable into a testable, observable outcome statement.

## Critical Rule — No Contradictions, No Invented Technical Specifics
- Cross-check every criterion against the MOM for consistency. If the MOM specifies a constraint (e.g., "batch processing, not real-time"), the criteria must respect that constraint exactly — never describe something as "real-time" if the MOM says batch/scheduled, and never claim a capability stronger than what was scoped.
- Do not invent specific technical methods/techniques (e.g., OCR, specific algorithms, specific AI models) unless the MOM explicitly names them. If the MOM only says "AI-driven extraction" or "automatically read using AI," keep the criterion at that same level of generality — do not add invented technical specificity.
- Before finalizing, mentally re-read each MOM constraint (timing, access model, region, scope limits) and verify no criterion contradicts or overstates it.

## How to phrase each bullet
Convert future-tense capability statements into verifiable, past/present-tense outcome statements, e.g.:
- Scope says "AI-driven processing will extract fields and apply business rules" → Acceptance criterion: "The system successfully extracts the predefined fields and applies the agreed business rules to generate the required output."
- Scope says "Users will access a web UI to upload documents" → Acceptance criterion: "Users are able to access the web interface and upload documents without errors."
Use outcome-oriented phrasing: "successfully established," "able to," "functioning as expected," "accessible," "generated successfully," "working as expected."

## Structure
- List one bullet per major capability/deliverable in the solution (derived from MOM/Scope) — cover the full solution flow: access/entry point, core processing/functionality, data handling, output/deliverable, and any infra-specific elements (connectivity, storage, sync) ONLY if they're part of this particular solution.
- The final bullet should almost always be a closing statement confirming the end-to-end workflow functions as expected and is jointly validated/accepted by Operisoft and the Customer.

## STRICT LENGTH RULE — count follows complexity, not a fixed number
- Do NOT force a fixed range like 8-12.
- A simple solution (e.g., single UI + AI processing + one output) should have roughly 4-6 criteria.
- A complex, multi-component solution (e.g., multi-location infra, multiple systems, multiple integrations) can have 8-12+ criteria — one per distinct verifiable component.
- Never pad with generic or repetitive criteria just to reach a higher count.

## Formatting rules
- Use "Ø  " (two spaces) for each bullet
- No markdown headers, no bold/italics
- Formal, third-person, concise — one line per criterion where possible
- Do NOT include pricing, timelines, or resourcing details
- Do NOT invent criteria for capabilities that aren't part of this solution

## Examples of expected style and depth

Example A (simple, application/AI solution — 5 criteria):
Ø  Users are able to access the web interface and upload PDF documents without errors.
Ø  The system successfully extracts the predefined fields from the uploaded PDF documents and applies the agreed business rules to generate the required derived data.
Ø  The solution accurately identifies and interprets relevant contextual information within the document to ensure consistent and structured output generation.
Ø  A structured Excel file containing the extracted fields and derived column is generated successfully and is available for download through the user interface.
Ø  The end-to-end workflow from document upload to Excel output generation functions as expected and is jointly validated and accepted by Operisoft and the Customer.

Example B (simple, chatbot solution — 5 criteria):
Ø  Users can access the sample chatbot UI and ask questions related to data stored in the Customer's NoSQL database.
Ø  The chatbot is able to understand user questions expressed in natural language and identify the required information.
Ø  The system retrieves relevant data from the NoSQL database using read-only access and provides clear, human-readable responses through the sample UI.
Ø  Chat interactions and responses are displayed correctly within the sample UI, demonstrating conversational continuity.
Ø  The end-to-end workflow functions as expected within the Customer's AWS account, with all required AWS services deployed and managed through AWS CloudFormation, and is jointly validated by Operisoft and the Customer.

Example C (complex, infrastructure/networking solution — 11 criteria):
Ø  Successful connectivity established between all three locations and AWS via Site-to-Site VPN
Ø  Remote users able to securely access AWS resources via Remote VPN
Ø  Amazon FSx accessible to all users with proper drive mapping
Ø  Individual user storage and shared storage functioning as expected
Ø  User authentication and access control working via AWS Managed Active Directory
Ø  Users able to read/write data based on assigned permissions
Ø  Amazon S3 bucket successfully storing production data
Ø  Automated data sync from production systems to S3 running twice daily without failure
Ø  Incremental backup and versioning working as expected
Ø  No critical performance or access issues observed during testing
Ø  Customer team successfully validated access, storage, and backup functionality

Example D (consistency example — batch, not real-time):
Correct: "Ø  The dashboard displays the status of all invoices processed in the most recent daily batch, including counts and breakdowns by status category."
Incorrect (do NOT do this — contradicts MOM's stated batch-only requirement): "Ø  A real-time dashboard view displays live invoice status as documents are processed."

## Output
Generate ONLY the acceptance criteria bullet points in the above style — no section heading, no preamble, no explanation. Let the number of bullets naturally reflect the complexity of this specific solution.
"""
    return call_bedrock(prompt)
