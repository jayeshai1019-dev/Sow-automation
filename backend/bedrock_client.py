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


def generate_project_objectives(customer_name: str, mom_text: str) -> str:
    prompt = f"""Based on the following MOM notes, write a professional 'Project Objectives and Overview' section for an AWS Cloud Migration Statement of Work.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Write 2-3 paragraphs covering the project background, what the customer wants to achieve, and the migration goals. Be professional and specific. Do not include headings."""
    return call_bedrock(prompt)


def generate_current_landscape(customer_name: str, mom_text: str) -> str:
    prompt = f"""Based on the following MOM notes, describe the customer's current IT landscape for an AWS SOW document.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Write 1-2 paragraphs covering their existing infrastructure, servers, applications, and pain points. Be professional. Do not include headings."""
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


def generate_tco_assumptions(customer_name: str, mom_text: str) -> str:
    prompt = f"""Based on the following MOM notes, write the 'Assumptions taken for cost calculation' for an AWS SOW.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Write 6-8 bullet points (starting with •) about cost assumptions including regions, compute sizing, backup, storage, and uptime. Be specific where information is available."""
    return call_bedrock(prompt)


def generate_high_level_scope(customer_name: str, mom_text: str) -> str:
    prompt = f"""Based on the following MOM notes, write the 'High Level Scope' section (3.1) for a Statement of Work.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Format the response as:
Ø  Operisoft will implement [main solution description], capable of:
·        [key capability 1]
·        [key capability 2]
·        [key capability 3]
·        [key capability 4]
·        [key capability 5]

Use the following as a reference template (adapt based on the customer requirements from MOM):
- Start with "Ø  Operisoft will implement..."
- List 5-7 key capabilities that describe what the solution will do
- Each capability should start with "·        " (bullet point with spacing)
- Be specific about AWS services and technologies mentioned in the MOM
- Focus on technical capabilities and deliverables

Example format:
Ø  Operisoft will implement a GenAI-based PDF Data Extraction and Automation Solution, capable of:
·        A web-based UI for uploading PDF documents, hosted on AWS Amplify.
·        Secure storage of uploaded PDF documents in Amazon S3 (landing bucket) for further processing.
·        Automated processing of uploaded PDFs using AWS Lambda functions to split, manage sessions, and orchestrate the extraction workflow.

Write ONLY the formatted text, no headings or additional explanations."""
    return call_bedrock(prompt)


def generate_scope_of_work(customer_name: str, mom_text: str) -> str:
    prompt = f"""Based on the following MOM notes, write the 'Scope of Work' section (3.2) for a Statement of Work.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Format the response with multiple paragraphs, each starting with "Ø  " covering:
1. Overall POC/project scope and what Operisoft will design/implement
2. Web interface description (if applicable)
3. Cloud infrastructure and AWS services integration
4. AI/processing capabilities (if applicable)
5. Deliverables and output
6. AWS services configuration and integration details
7. Validation and testing approach

Then add a "Notes:" section with 3-4 bullet points starting with "Ø  " covering:
- POC accuracy disclaimer (if applicable)
- Assumptions and future service needs
- Production SOW requirements

Use the following as a reference template (adapt based on customer requirements from MOM):
Ø  As part of the POC, Operisoft will design and implement...
Ø  A web-based user interface will be developed...
Ø  Uploaded documents/data will be securely stored...
Ø  AI-driven processing using...
Ø  The final output will be...
Ø  For the purpose of this POC, Operisoft will configure and integrate...
Ø  During the POC phase, Operisoft and the Customer team will collaboratively...

Notes:
Ø  As this is a POC project...
Ø  This solution is based on certain assumptions...
Ø  For production-level implementation...

Write ONLY the formatted text with proper bullet structure (Ø), no section headings."""
    return call_bedrock(prompt)


def generate_acceptance_criteria(customer_name: str, mom_text: str) -> str:
    prompt = f"""Based on the following MOM notes, write the 'Acceptance Criteria' section (3.4) for a Statement of Work.
Customer Name: {customer_name}
MOM Notes: {mom_text}

Generate 8-12 acceptance criteria bullet points that define what must be successfully demonstrated for project acceptance.
Each bullet point should start with "Ø  " and describe a specific, measurable criterion.

Cover these aspects based on what's relevant from the MOM:
1. Connectivity and network access (VPN, networking)
2. User access and authentication
3. Storage and file system functionality
4. Data processing and extraction capabilities
5. Integration with required systems
6. Backup and data sync operations
7. Performance and reliability
8. User validation and final acceptance

Use the following as reference templates (adapt based on customer requirements from MOM):

Example 1 (Infrastructure/Networking):
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

Example 2 (Application/Data Processing):
Ø  Users are able to access the web interface and upload PDF documents without errors.
Ø  The system successfully extracts the predefined fields from the uploaded PDF documents and applies the agreed business rules to generate the required derived data.
Ø  The solution accurately identifies and interprets relevant contextual information within the document to ensure consistent and structured output generation.
Ø  A structured Excel file containing the extracted fields and derived column is generated successfully and is available for download through the user interface.
Ø  The end-to-end workflow from document upload to Excel output generation functions as expected and is jointly validated and accepted by Operisoft and the Customer.

Write 8-12 specific acceptance criteria starting with "Ø  " that are relevant to the customer's project requirements.
Write ONLY the bullet points, no headings or additional explanations."""
    return call_bedrock(prompt)
