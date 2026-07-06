import os
import re
import tempfile
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

from bedrock_client import (
    generate_about_customer,
    generate_project_objectives,
    generate_current_landscape,
    generate_key_highlights,
    generate_dr_requirements,
    generate_proposed_drs_solution,
    generate_tco_assumptions,
    generate_high_level_scope,
    generate_scope_of_work,
    generate_acceptance_criteria,
)
from pricing_client import build_pricing_estimate

# ── helpers ──────────────────────────────────────────────────────────────────

ORANGE = RGBColor(0xFF, 0x6B, 0x00)   # Operisoft brand orange
DARK   = RGBColor(0x1F, 0x27, 0x63)   # dark navy

def _set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def _heading(doc, text, level=1, color=None):
    p = doc.add_paragraph()
    p.style = doc.styles["Heading %d" % level] if ("Heading %d" % level) in [s.name for s in doc.styles] else doc.styles["Normal"]
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14 if level == 1 else 12 if level == 2 else 11)
    run.font.color.rgb = color or DARK
    return p

def _para(doc, text, bold=False, italic=False, size=10, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return p

def _bullet(doc, text, size=10):
    p = doc.add_paragraph(style="List Bullet") if "List Bullet" in [s.name for s in doc.styles] else doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    run = p.add_run(("• " if p.style.name != "List Bullet" else "") + text)
    run.font.size = Pt(size)
    return p

def _add_bullets_from_text(doc, text):
    """Parse bullet text (lines starting with • or -) and add as bullet paras."""
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        line = line.lstrip("•-·").strip()
        if line:
            _bullet(doc, line)

def _section_divider(doc):
    doc.add_paragraph()

def _two_col_table(doc, rows_data, header=None):
    """Create a simple 2-column table."""
    cols = 2
    table = doc.add_table(rows=len(rows_data) + (1 if header else 0), cols=cols)
    table.style = "Table Grid"
    if header:
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(header):
            hdr_cells[i].text = h
            hdr_cells[i].paragraphs[0].runs[0].bold = True
            _set_cell_bg(hdr_cells[i], "1F2763")
            hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for idx, (k, v) in enumerate(rows_data):
        row = table.rows[idx + (1 if header else 0)]
        row.cells[0].text = k
        row.cells[1].text = v
    return table


# ── static content blocks ──────────────────────────────────────────────────

CONFIDENTIALITY_TEXT = (
    "The information contained in this document has been prepared to be used in the context "
    "of this project. It should not be used as a model or precedent in any situation outside "
    "of this project. This document must not be copied or reproduced by any means without the "
    "authorization of the parties involved. A good effort has gone into the preparation of this "
    "document to ensure that the information presented is correct at the time of printing. The "
    "parties involved in this project assume no responsibility for any errors that may arise in "
    "the application of this information in a context other than the project for which it was prepared."
)

ABOUT_OPERISOFT = (
    "Operisoft Technologies Private Limited is a leading software consulting and services provider "
    "in digital transformation and Cloud Technologies, headquartered in Mumbai with offices in Pune, "
    "Bangalore, Hyderabad and Ahmedabad. We specialize in cloud technologies, including Amazon Web "
    "Services (AWS), providing expert support for complex solutions and the latest cloud innovations. "
    "With a team of over 55+ skilled professionals across India, we deliver best-in-class solutions "
    "tailored to customer needs.\n\n"
    "Operisoft supports clients across diverse sectors—including manufacturing, pharmaceuticals, "
    "healthcare, automotive, plant and machinery, chemical and process industries, e-commerce and "
    "retail, media and entertainment, as well as education and non-profit organizations—by leveraging "
    "cloud technologies like AWS and SaaS platforms. We proudly serve over 200 customers across these industries.\n\n"
    "With over 5+ years of experience as a leading AWS Cloud Partner, Operisoft excels in guiding "
    "businesses through cloud migration and modernization. Our team boasts the highest credentials, "
    "including 3 major AWS competencies & service validations, and certified AWS specialists, with "
    "solution architects holding 20+ active AWS certifications. As an AWS Advanced Tier Services Partner, "
    "we showcase excellence in AWS Managed Services, Data & Analytics, Migration, and DevOps. We were "
    "honoured as the Rising Star Partner of the Year 2023 in India."
)

TESTING_MONITORING_TEXT = [
    "Testing of the application will be done by client after the deployment.",
    "Once the Testing of the application is successful, the DNS records will be updated for going live.",
]

MONITORING_INFRA_BULLETS = [
    "Detailed monitoring will be done for all servers using AWS CloudWatch.",
    "For monitoring of OS level metrics like memory and disk utilization, custom CloudWatch metrics to be used.",
    "We will configure CloudWatch agent for custom CloudWatch monitoring.",
    "CloudWatch Alerts will be configured if resource utilization on any of the servers crosses a specific threshold.",
    "Monitoring parameters and thresholds will be configured as per the customer KPIs.",
    "Alerts will be configured for the status of the server for any issue on the server and server is not ready status.",
]

EDR_TEXT = (
    "AWS EDR (Elastic Disaster Recovery) minimizes downtime and data loss with fast, reliable recovery "
    "of on-premises and cloud-based applications using affordable storage, minimal compute, and "
    "point-in-time recovery.\n\nDisaster Recovery utilizes block-level, Continuous Data Replication, "
    "which ensures that target machines are spun up in their most up-to-date state during a disaster "
    "or drill. Organizations can thereby achieve sub-second Recovery Point Objectives (RPOs)."
)

DRS_ARCH_BULLETS = [
    "Replication server status reporting",
    "Staging area resources automatically created and terminated",
    "Recovery instances launched with RTO of minutes and RPO of few Minutes",
    "Continuous block-level replication (compressed and encrypted)",
]

IAM_BULLETS = [
    "The root account will not be used for day-to-day activities and the credentials of the root account will remain with the customer team only.",
    "AWS console will be accessed through cross-account IAM roles. We recommend using SAML-based federated users to access the AWS account instead of IAM users.",
    "MFA will be enabled for root AWS accounts, and we suggest and strongly recommend using MFA for other IAM users too.",
    "AWS partner users will use cross-account IAM roles to access the AWS account and there will be no IAM user for the AWS partner team.",
    "A strong password policy will be in place which includes regular password rotation.",
    "Password length should not be less than 8 characters.",
    "Password should contain at least one character from a-z, A-Z, 0-9, and special characters.",
    "Old passwords cannot be repeated.",
    "A role-based access policy with the principle of least privilege will be configured for accessing the AWS console and other maintenance tasks.",
]

IAM_ANALYZER_BULLETS = [
    "AWS IAM access analyzer helps you to fine-grain the continued cycle of access to the least privileged.",
    "AWS IAM Access Analyzer helps you streamline permissions management throughout each step of the cycle.",
    "Access Analyzer will identify the resources shared with external principles by using logic-based reasoning to analyze resource-based policy in your AWS environment.",
    "Access analyzer can help you to validate your created policy against IAM best practices.",
    "Access analyzer will be able to analyze your AWS CloudTrail log to identify actions and services that have been used by an IAM entity.",
]

DETECTIVE_CONTROLS_BULLETS = [
    "CloudTrail Logs will be enabled for all users. The last 7 days' data can be monitored from AWS CloudTrail Console.",
    "VPC Flow logs will be enabled, and Guard Duty will be configured to monitor VPC Flow Logs.",
    "GuardDuty will generate findings whenever it detects unexpected and potentially malicious activity in your AWS environment.",
    "CloudTrail Logs will be stored in a private S3 bucket. Only permitted users will have access permission to that bucket.",
    "We recommend pushing access and other logs from all servers to CloudWatch so that logs can be analyzed.",
    "All the sys logs and application logs will be stored in S3 for 90 days.",
    "After 90 days, logs will be moved from S3 to Glacier.",
    "Logs in Glacier will be retained for another 90 days before they are permanently removed.",
    "Logs stored on S3 can be analyzed using AWS Athena.",
    "AWS config rules will be configured to monitor Config logs stored on the s3 bucket and used for dynamic compliance checking.",
    "CloudTrail Logs will be streamed to AWS CloudWatch and filters will be created in CloudWatch for modification in security groups, start-stop of EC2 instances, or use of root user APIs.",
]

AWS_DETECTIVE_BULLETS = [
    "Amazon Detective will enable you to analyze and visualize security data from your AWS CloudTrail logs, VPC Flow logs, and Amazon GuardDuty findings.",
    "Amazon Detective will be integrated with Amazon GuardDuty and AWS Security Hub as well as AWS partner security products.",
    "AWS Detective will provide an interactive unified view which will help you to visualize all the context and details in one place.",
]

SECURITY_HUB_BULLETS = [
    "AWS Security Hub will centralize and prioritize security findings from across AWS accounts, services, and supported third-party partners to help you analyze your security trends and identify the highest priority security issues.",
    "Security Hub will help you manage security across multiple AWS accounts within a region by configuring the multi-account hierarchy within AWS Security Hub.",
    "Security Hub will be deployed with CloudFormation Template.",
    "Security Hub can be integrated with AWS Detective, Guard Duty, and other security services.",
    "Creation and Configuration of security events alerts from AWS Detective.",
]

PARTNER_DELIVERABLES_BULLETS = [
    "Inputs provided from the client in any terms will be assumed to be accurate.",
    "AWS usage and billing will be borne by client during the development as well as at UAT and deployment.",
    "Necessary approvals and timely review of submissions.",
    "If a need of any new plugin etc. or any 3rd party service provider needed – will be arranged by client.",
    "Customer will complete testing of all the deployed AWS services before handover.",
    "Operisoft will provide consultation for AWS services design and deployment strategies and execute mutually agreed upon new change order.",
    "Operisoft will not be responsible for any technology or product limitations hindering the service delivery, support or timelines.",
]

CUSTOMER_DEPS_BULLETS = [
    "The customer team will provide the details of network architecture and application dependencies.",
    "Availability of the key/right stakeholders will be assured for defined activities MRA/Discovery Session.",
    "Data migration is the key to success for this project and it is largely dependent on adequate network bandwidth availability. The customer will have to ensure the availability of sufficient network bandwidth at the source data centre.",
    "Bandwidth requirements will be mutually agreed upon during the initial phase of the feasibility of migration methodology.",
    "All application-level configuration changes to be done by the customer in-house team and AWS PARTNER will provide support related to AWS-specific configurations.",
    "All third-party tools and services will be migrated as-is from the current data centre.",
    "Licenses for all applications and third-party tools will be provided by the customer.",
    "Installation of agents for application migration service will be done by the customer's IT team.",
    "Any changes required in application code/configuration or database configuration will be done by the customer's Team.",
    "The customer team has to share the Applications dependencies along with compatibility with Operating System licenses etc.",
    "Providing access control details for IAM, S3, Security Groups as per business requirement.",
    "Customer needs to align their IT and network teams for setting up the network architecture between AWS and their on-premises environment.",
]

ASSUMPTIONS_BULLETS = [
    "A single point of contact will be assigned from the customer team end for coordination and providing the required information, access, approvals, and sign-off.",
    "All required information and approvals from the customer team will be provided within 3 working days.",
    "The customer will provide the complete know-how on the application deployment architecture, access, and data flows to Operisoft team.",
    "The customer will provide required access to their environments for assessment for AWS migration.",
    "Account structure and network architecture will be finalized with customer's IT teams.",
    "Compute sizing is designed based on the utilization report and existing server configuration.",
    "Inputs provided from the client in any terms will be assumed to be accurate.",
    "AWS usage and billing will be borne by client during the development as well as at UAT and deployment.",
    "Necessary approvals and timely review of submissions.",
    "Customer will complete testing of all the deployed AWS services before handover.",
    "Operisoft will not be responsible for any technology or product limitations hindering the service delivery, support or timelines.",
    "Operisoft's timely and adequate performance of the services shall depend upon full access to appropriate customer personnel and information.",
    "In case of sudden spikes in volumes due to unforeseen circumstances, the support will be on a best effort basis.",
    "All the organizational or technical changes which may affect any services will be communicated in writing to Operisoft at least 1 week prior to changes.",
    "We assume for business-critical applications necessary HA architecture in place to meet the requirements.",
    "Unauthorized changes: In the event of unauthorized changes in AWS infrastructure, Operisoft reserves rights to refuse support.",
]

EXCLUSIONS_BULLETS = [
    "Security configurations and settings outside of the AWS environment will not be part of the scope.",
    "Security configuration of Existing environment or in the corporate network environment.",
    "Any change in OS, application, or third-party tool/service configuration.",
    "Changes in application code or configuration in the application.",
    "3rd party App Integrations is not included in this scope.",
    "Any activities related to modernization will be done after the migration phase.",
    "Any new cloud solution build and deployment requirement including but not limited to application modernization will be a separate engagement.",
    "Any change in the current On-Premises environment.",
    "Application testing.",
]

RISK_TABLE_DATA = [
    ("1", "The large size of the database may cause longer downtime", "Test the feasibility of database migration with all available options."),
    ("2", "As the project involves migration to the cloud along with changes in the application configuration and architecture might be required", "This would require higher involvement of the application team."),
    ("3", "Network bandwidth constraint in uploading data/server replication", "The client team will ensure enough bandwidth for migration activity."),
    ("4", "Availability of resources and teams", "Stand-by team will be employed for the project activities."),
    ("5", "Coordination among multiple teams", "SPOC will be identified from each team for the project activity."),
    ("6", "Non-availability of information", "Checklists will be prepared for all required information."),
    ("7", "Delay in approvals and sign-off", "Regular reviews will be done to check the status of the project tasks."),
]

PROJECT_PLAN_TASKS = [
    ("1", "Contract Sign Off"),
    ("", "Project Initiation / Requirement Gathering"),
    ("2", "Finalization of ownership of activities between Operisoft and the Customer"),
    ("3", "Prepare and share Detailed Project Plan"),
    ("4", "Coordination with application team for Knowledge Transfer"),
    ("5", "Create New AWS Account for the Customer"),
    ("6", "Sign off on Documentation by the Customer"),
    ("", "Setting up Landing Zone"),
    ("7", "Create AWS Control Tower and Landing zone (Master, Security and Log Accounts)"),
    ("8", "Create AWS account and network design"),
    ("9", "Create OU for Network and Shared services account"),
    ("10", "Configure Network design review and Access control"),
    ("11", "Creation of Network, shared services, Prod and Dev/test workloads"),
    ("12", "Setting up the control tower and preparation of Cloud Formation template for Network"),
    ("13", "AWS Account Setup"),
    ("", "Data Migration"),
    ("14", "Setup MGN Service for Data Migration"),
    ("15", "Migrate the data from existing Datacenter site to proposed AWS Cloud"),
    ("16", "Test data integrity and validate data consistency"),
    ("", "Setup Development Environment"),
    ("17", "Set up VPC, Subnets, NACLs, NAT Gateway and Security Groups"),
    ("18", "Deploy Application EC2 instances in private subnet as per shared sizing"),
    ("19", "Set up an Application Load Balancer in the private subnet"),
    ("20", "Deploy Database EC2 instances in Private subnet"),
    ("21", "Test Application Functionality"),
    ("", "Setup UAT Environment"),
    ("22", "Set up VPC, Subnets, NACLs, NAT Gateway and Security Groups"),
    ("23", "Deploy Application EC2 instances in private subnet as per shared sizing"),
    ("24", "Set up an Application Load Balancer in the private subnet"),
    ("25", "Deploy Database EC2 instances in Private subnet"),
    ("26", "Test Application Functionality"),
    ("", "Setup Production Environment"),
    ("27", "Set up VPC, Subnets, NACLs, NAT Gateway and Security Groups"),
    ("28", "Deploy EC2 instances in private subnet in multi-AZ Mode"),
    ("29", "Set up an Application Load Balancer in the private subnet"),
    ("30", "Deploy Database EC2 instances in Private subnet"),
    ("31", "Test Application Functionality"),
    ("", "Setup Disaster Recovery"),
    ("32", "Set up VPC, Subnets, NACLs, NAT Gateway and Security Groups"),
    ("33", "Set up AWS DRS Service to replicate Application and Database Servers to DR region"),
    ("34", "Test Replication and DR Setup"),
    ("", "Implementation & Review Of AWS Best Security Practices"),
    ("35", "Configuration of Password policy in IAM in all accounts"),
    ("36", "Setup AWS Route53 and update DNS records"),
    ("37", "Setup and configuration of AWS backup snapshot policies"),
    ("38", "Setup File System and database level Backup"),
    ("39", "Enable CloudTrail to keep a track of account level activities"),
    ("40", "Configure CloudWatch dashboards along with alerts"),
    ("41", "Configure AWS GuardDuty to detect threats and generate alerts"),
    ("42", "Configure AWS Security HUB"),
    ("43", "AWS KMS setup for data encryption"),
    ("44", "Install Antivirus"),
    ("45", "Setup Palo Alto Firewall"),
    ("46", "Setup AWS WAF"),
    ("47", "Setup AWS Inspector"),
    ("", "Testing, Validation and Support"),
    ("48", "Handover of Document and SOP post implementation of setup"),
    ("49", "Testing of and troubleshooting of application functionality on the Production Site"),
    ("50", "Validation & Sign-Off by the Customer"),
]

COMMERCIAL_TERMS = """i. Offer Validity: Upto 30 days.

ii. Payment Term:
   a. Implementation: 100% advance unless explicitly mentioned.
   b. Managed Services/AWS Bills: 15 days after Invoice generation.

iii. Taxes: Extra @ 18% IGST/CGST+SGST in case of INR billing

iv. SAC/HSN:
   a. 998315 Cloud consumption & services
   b. 997331 Software

v. Billing Cycle: Monthly invoices are generated from 5th to 15th of the next month. And must be paid within 15 days to avoid stopping the services.

vi. Order to be placed on:
OPERISOFT TECHNOLOGIES PRIVATE LIMITED
Office No 301, Plot No 273, Zion Complex Building
Sector-10, Khaghar, Navi Mumbai – 410210"""


LANDING_ZONE_MANDATORY_GUARDRAILS = [
    "Disallow Changes to Encryption Configuration for AWS Control Tower Created Amazon S3 Buckets in Log Archive",
    "Disallow Changes to Logging Configuration for AWS Control Tower Created Amazon S3 Buckets in Log Archive",
    "Disallow Changes to Bucket Policy for AWS Control Tower Created Amazon S3 Buckets in Log Archive",
    "Disallow Changes to Lifecycle Configuration for AWS Control Tower Created Amazon S3 Buckets in Log Archive",
    "Disallow Changes to Amazon CloudWatch Logs Log Groups set up by AWS Control Tower",
    "Disallow Deletion of AWS Config Aggregation Authorizations Created by AWS Control Tower",
    "Disallow Deletion of Log Archive",
    "Detect Public Read Access Setting for Log Archive",
    "Detect Public Write Access Setting for Log Archive",
    "Disallow Configuration Changes to CloudTrail",
    "Integrate CloudTrail Events with Amazon CloudWatch Logs",
    "Enable CloudTrail in All Available Regions",
    "Enable Integrity Validation for CloudTrail Log File",
    "Disallow Changes to Amazon CloudWatch Set Up by AWS Control Tower",
    "Disallow Changes to Tags Created by AWS Control Tower for AWS Config Resources",
    "Disallow Configuration Changes to AWS Config",
    "Enable AWS Config in All Available Regions",
    "Disallow Changes to AWS Config Rules Set Up by AWS Control Tower",
    "Disallow Changes to AWS IAM Roles Set Up by AWS Control Tower and AWS CloudFormation",
    "Disallow Changes to AWS Lambda Functions Set Up by AWS Control Tower",
    "Disallow Changes to Amazon SNS Set Up by AWS Control Tower",
    "Disallow Changes to Amazon SNS Subscriptions Set Up by AWS Control Tower",
]

LANDING_ZONE_RECOMMENDED_GUARDRAILS = [
    "Disallow the Creation of Access Keys for the Root User",
    "Disallow Actions as a Root User",
    "Detect Whether Encryption is Enabled for Amazon EBS Volumes Attached to Amazon EC2 Instances",
    "Detect Whether Unrestricted Incoming TCP Traffic is Allowed",
    "Detect Whether Unrestricted Internet Connection Through SSH is Allowed",
    "Detect Whether MFA for the Root User is Enabled",
    "Detect Whether Public Read Access to Amazon S3 Buckets is Allowed",
    "Detect Whether Public Write Access to Amazon S3 Buckets is Allowed",
    "Detect Whether Amazon EBS Volumes are Attached to Amazon EC2 Instances",
    "Detect Whether Amazon EBS Optimization is Enabled for Amazon EC2 Instances",
    "Detect Whether Public Access to Amazon RDS Database Instances is Enabled",
    "Detect Whether Public Access to Amazon RDS Database Snapshots is Enabled",
    "Detect Whether Storage Encryption is Enabled for Amazon RDS Database Instances",
    "Enable Encryption at Rest for Log Archive",
    "Enable Access Logging for Log Archive",
    "Disallow Policy Changes to Log Archive",
    "Set a Retention Policy for the Log Archive",
    "Disallow Changes to Replication Configuration for Amazon S3 Buckets",
    "Disallow Delete Actions on CloudTrail S3 Buckets Without MFA",
    "Detect Whether MFA is Enabled for AWS IAM Users",
    "Detect Whether MFA is Enabled for AWS IAM Users of the AWS Console",
    "Detect Whether Versioning for Amazon S3 Buckets is Enabled",
]

PALOALTO_BULLETS = [
    "Application visibility for informed security decisions: The VM-Series provides application visibility across all ports.",
    "Segment/Whitelist applications for security and compliance: Using segmentation and whitelisting policies allows you to control applications communicating across different subnets.",
    "Prevent advanced attacks within allowed application flows: The VM-Series allows you to use Palo Alto Networks Threat Prevention, DNS Security, and WildFire® to apply application-specific policies.",
    "Control application access with user-based policies: Integration with Active Directory®, LDAP, and other user repositories.",
    "Policy consistency through centralized management: Panorama™ provides centralized network security management for your VM-Series firewalls.",
    "Cloud-native scalability and availability: In public cloud environments, we recommended using cloud services such as application gateways, load balancers, and automation.",
]

MGN_BULLETS = [
    "AWS partner team will set up the network infrastructure - VPN from Onprem/Other to AWS for secure replication.",
    "AWS partner team will create a separate VPC environment for isolation and security.",
    "Deployment of AWS MGN tool in Customers existing setup to initiate the replication of server data.",
    "Setup of Staging VPC and replication VM in AWS with EBS volume.",
    "AWS MGN is a highly automated lift-and-shift solution, which works by replicating (physical or virtual) and/or cloud servers into assigned AWS account.",
    "First time setup: Creating the replication template by initializing the AWS Application Migration Service.",
    "Adding source servers: Install the AWS Replication Agent on both Linux and Windows source servers.",
    "Configuring launch settings: Configure instructions that determine how a test or cutover instance will be launched.",
    "Launching a test instance: Test the migration of source servers to AWS prior to initiating a cutover.",
    "Launching a cutover instance: Once testing is finalized, the cutover will migrate source servers to AWS.",
]


# ── main document builder ─────────────────────────────────────────────────

def generate_sow_document(data: dict) -> str:
    customer_name         = data.get("customer_name", "Customer")
    mom_text              = data.get("mom_text", "")
    company_url           = data.get("company_url", "")
    include_landing_zone  = data.get("include_landing_zone", False)
    include_control_tower = data.get("include_control_tower", False)
    include_lz_arch       = data.get("include_landing_zone_arch", False)
    include_paloalto      = data.get("include_paloalto", False)
    include_mgn           = data.get("include_mgn_migration", False)

    # Extract company info from URL if provided
    company_about_us = ""
    if company_url and company_url.strip():
        from url_extractor import extract_company_info
        extracted = extract_company_info(company_url.strip())
        company_about_us = extracted.get("about_us", "")
        # Also use extracted company name if customer_name wasn't explicitly set
        if not customer_name or customer_name == "Customer":
            customer_name = extracted.get("company_name", customer_name)

    # Generate dynamic content via Bedrock (credentials loaded from .env)
    about_customer      = generate_about_customer(customer_name, mom_text, company_about_us)
    proj_objectives     = generate_project_objectives(customer_name, mom_text)
    current_landscape   = generate_current_landscape(customer_name, mom_text)
    key_highlights      = generate_key_highlights(customer_name, mom_text)
    dr_reqs             = generate_dr_requirements(customer_name, mom_text)
    proposed_drs        = generate_proposed_drs_solution(customer_name, mom_text)
    tco_assumptions     = generate_tco_assumptions(customer_name, mom_text)
    high_level_scope    = generate_high_level_scope(customer_name, mom_text)
    scope_of_work       = generate_scope_of_work(customer_name, mom_text)
    acceptance_criteria = generate_acceptance_criteria(customer_name, mom_text)

    # Build real AWS Pricing Calculator estimate via MCP
    pricing             = build_pricing_estimate(customer_name, mom_text, key_highlights)

    doc = Document()

    # ── Page margins ──
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── Cover Page ────────────────────────────────────────────────────────
    doc.add_paragraph()
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run("Proposal for AWS Cloud Services")
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = DARK

    doc.add_paragraph()
    for_p = doc.add_paragraph()
    for_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for_p.add_run("For").font.size = Pt(14)

    doc.add_paragraph()
    cust_p = doc.add_paragraph()
    cust_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = cust_p.add_run(customer_name)
    run2.bold = True
    run2.font.size = Pt(18)
    run2.font.color.rgb = ORANGE

    doc.add_paragraph()
    doc.add_paragraph()

    # Document History table
    hist_table = doc.add_table(rows=2, cols=3)
    hist_table.style = "Table Grid"
    headers = ["Version", "Date", "Submitted by"]
    for i, h in enumerate(headers):
        cell = hist_table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        _set_cell_bg(cell, "1F2763")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    hist_table.rows[1].cells[0].text = "1.0"
    hist_table.rows[1].cells[1].text = "Date"
    hist_table.rows[1].cells[2].text = "Operisoft Technologies"

    doc.add_page_break()

    # ── 1. Executive Summary ──────────────────────────────────────────────
    _heading(doc, "1. Executive Summary", 1, DARK)

    _heading(doc, "1.1 Confidentiality Notice", 2, ORANGE)
    _para(doc, CONFIDENTIALITY_TEXT)
    _section_divider(doc)

    _heading(doc, "1.2 About Operisoft", 2, ORANGE)
    for para in ABOUT_OPERISOFT.split("\n\n"):
        _para(doc, para.strip())
    _section_divider(doc)

    _heading(doc, "1.3 About Customer", 2, ORANGE)
    _para(doc, about_customer)
    _section_divider(doc)

    # ── 2. Project Details ────────────────────────────────────────────────
    _heading(doc, "2. Project Details", 1, DARK)

    _heading(doc, "2.1 Project Objectives and Overview", 2, ORANGE)
    for para in proj_objectives.split("\n\n"):
        if para.strip():
            _para(doc, para.strip())
    _section_divider(doc)

    _heading(doc, "2.2 Understanding of Current Landscape", 2, ORANGE)
    for para in current_landscape.split("\n\n"):
        if para.strip():
            _para(doc, para.strip())
    _section_divider(doc)

    # ── 3. Proposed Solution ──────────────────────────────────────────────
    _heading(doc, "3. Proposed Solution", 1, DARK)

    _heading(doc, "3.1 High Level Scope", 2, ORANGE)
    _para(doc, high_level_scope)
    _section_divider(doc)

    _heading(doc, "3.2 Scope of Work", 2, ORANGE)
    _para(doc, scope_of_work)
    _section_divider(doc)

    _heading(doc, "3.3 Key Highlights of the Solution", 2, ORANGE)
    _add_bullets_from_text(doc, key_highlights)
    _section_divider(doc)

    _heading(doc, "3.4 Acceptance Criteria", 2, ORANGE)
    _para(doc, acceptance_criteria)
    _section_divider(doc)

    _heading(doc, "3.5 Proposed Solution Diagram", 2, ORANGE)
    _para(doc, "[Proposed Solution Architecture Diagram – Please insert diagram here]", italic=True, color=RGBColor(0x99, 0x99, 0x99))
    _section_divider(doc)

    # Optional sections 3.6, 3.7, 3.8, 3.9
    sec_num = 6
    if include_landing_zone:
        _heading(doc, "3.6 Implementation of Landing Zone", 2, ORANGE)
        _para(doc, "In this solution, we propose setting up a Landing zone on the newly created AWS account setup. We are proposing the following Landing Zone architecture:")
        doc.add_paragraph()
        lz_table = doc.add_table(rows=1, cols=3)
        lz_table.style = "Table Grid"
        for i, h in enumerate(["Organization Root", "Organization Unit", "Account"]):
            cell = lz_table.rows[0].cells[i]
            cell.text = h
            cell.paragraphs[0].runs[0].bold = True
            _set_cell_bg(cell, "1F2763")
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        ou_data = [
            ("Organizational Root Account", "Master", "Management"),
            ("", "Security (OU)", "Security"),
            ("", "", "Log Archive"),
            ("", "Infra (OU)", "Network"),
            ("", "", "Observability"),
            ("", "", "Shared Services"),
            ("", "Sandbox (OU)", "Account for UAT and Dev application deployment"),
            ("", "Workload (OU)", "Account for Production application deployment"),
        ]
        for row_data in ou_data:
            row = lz_table.add_row()
            for i, val in enumerate(row_data):
                row.cells[i].text = val
        _section_divider(doc)
        _para(doc, "There will be a single master account for the entire organization. Under the organization's root, there will be a central security OU and separate OUs for each business unit/LOB.")
        _para(doc, "Mandatory Guardrails:", bold=True)
        for g in LANDING_ZONE_MANDATORY_GUARDRAILS:
            _bullet(doc, g)
        _para(doc, "Recommended Guardrails:", bold=True)
        for g in LANDING_ZONE_RECOMMENDED_GUARDRAILS:
            _bullet(doc, g)
        _section_divider(doc)
        sec_num += 1

    if include_control_tower:
        _heading(doc, "3.7 Configuration of Control Tower Setup", 2, ORANGE)
        ct_bullets = [
            "Enable Control Tower from AWS Console",
            "Register OU to Control Tower using the console",
            "Creation of a new Account in the Organization using Automation",
            "Configuration of service control policies using control tower customization automation",
            "Configuration of security Hub using CloudFormation",
            "Configuration of GuardDuty using CloudFormation",
        ]
        for b in ct_bullets:
            _bullet(doc, b)
        _section_divider(doc)

    if include_lz_arch:
        _heading(doc, "3.8 AWS Landing Zone Architecture", 2, ORANGE)
        lza_bullets = [
            "End user account provisioning through AWS Service Catalog (Account factory)",
            "Centralized monitoring and notifications for each OU using Amazon CloudWatch and Amazon SNS",
            "Centralized logging and immutable archive for all Cloud infrastructure and application logs",
            "Workload OU and Accounts to be defined for separate Business Units, applications, or environments",
            "Isolation of resources using separate accounts for similar groups of workloads",
            "Automation for compliance management and security posture management",
            "Automation for Organization policies like SCP, backup, and Tagging policies",
            "Automatic Tagging of resources for Billing segmentation for CloudFinOps",
        ]
        for b in lza_bullets:
            _bullet(doc, b)
        _section_divider(doc)

    if include_paloalto:
        _heading(doc, "3.9 PaloAlto Next Generation Firewall", 2, ORANGE)
        _para(doc, "We propose PaloAlto Next Generation Firewall with advanced security features like IPS, URL filtering, Bot Protection, Sandboxing which will inspect all the incoming and outgoing traffic.")
        _para(doc, "Key VM-Series Features and Capabilities:", bold=True)
        for b in PALOALTO_BULLETS:
            _bullet(doc, b)
        _section_divider(doc)

    # ── 4/5. Migration Approach ───────────────────────────────────────────
    _heading(doc, "4. Migration Approach", 1, DARK)

    if include_mgn:
        _heading(doc, "4.1 Migration of Data Using MGN", 2, ORANGE)
        for b in MGN_BULLETS:
            _bullet(doc, b)
        _section_divider(doc)

    _heading(doc, "4.2 Testing and Monitoring", 2, ORANGE)
    for t in TESTING_MONITORING_TEXT:
        _bullet(doc, t)
    _section_divider(doc)

    # ── 5/6. Monitoring ───────────────────────────────────────────────────
    _heading(doc, "5. Monitoring AWS Infrastructure", 1, DARK)
    for b in MONITORING_INFRA_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    # ── 6/7. DR ───────────────────────────────────────────────────────────
    _heading(doc, "6. Proposed DR Approach", 1, DARK)

    _heading(doc, "6.1 Requirement of DR", 2, ORANGE)
    dr_rows = [
        ("Identify Critical Systems", str(dr_reqs.get("critical_systems", "All production systems"))),
        ("Primary Site Location",     str(dr_reqs.get("primary_site", "Mumbai Region"))),
        ("Type of DR",                str(dr_reqs.get("dr_type", "Passive"))),
        ("Expected RTO",              str(dr_reqs.get("rto", "30 minutes"))),
        ("Expected RPO",              str(dr_reqs.get("rpo", "15 minutes"))),
    ]
    _two_col_table(doc, dr_rows)
    _section_divider(doc)

    _heading(doc, "6.2 AWS Elastic Disaster Recovery for Replication", 2, ORANGE)
    for para in EDR_TEXT.split("\n\n"):
        _para(doc, para.strip())
    _section_divider(doc)

    _heading(doc, "6.3 AWS DRS Architecture Design", 2, ORANGE)
    _para(doc, "[AWS DRS Architecture Diagram – Please insert diagram here]", italic=True, color=RGBColor(0x99, 0x99, 0x99))
    for b in DRS_ARCH_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    _heading(doc, "6.4 Proposed DRS Solution", 2, ORANGE)
    _add_bullets_from_text(doc, proposed_drs)
    _section_divider(doc)

    # ── 7. Cost Estimate ──────────────────────────────────────────────────
    _heading(doc, "7. Cost Estimate / Commercial Terms", 1, DARK)

    _heading(doc, "7.1 AWS Total Cost of Ownership [TCO]", 2, ORANGE)

    calc_url = pricing.get("url", "https://calculator.aws/pricing/2/estimate")

    link_para = doc.add_paragraph()
    link_para.paragraph_format.space_before = Pt(4)
    link_para.paragraph_format.space_after  = Pt(4)
    label_run = link_para.add_run("AWS Pricing Calculator Estimate: ")
    label_run.bold = True
    label_run.font.size = Pt(11)
    label_run.font.color.rgb = DARK
    url_run = link_para.add_run(calc_url)
    url_run.font.size = Pt(11)
    url_run.font.color.rgb = RGBColor(0x00, 0x56, 0xD2)
    url_run.underline = True

    _section_divider(doc)

    _heading(doc, "7.2 Operisoft Service Charges", 2, ORANGE)
    cost_rows = [
        ("Cost Component", "Cost in INR"),
        ("OTC [One Time Charges]", ""),
        ("Retainership cost [Recurring Managed services]", ""),
    ]
    svc_table = doc.add_table(rows=3, cols=2)
    svc_table.style = "Table Grid"
    for i, (k, v) in enumerate(cost_rows):
        svc_table.rows[i].cells[0].text = k
        svc_table.rows[i].cells[1].text = v
        if i == 0:
            svc_table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
            svc_table.rows[i].cells[1].paragraphs[0].runs[0].bold = True
    _para(doc, "GST Extra as applicable", italic=True)
    _section_divider(doc)

    _heading(doc, "7.3 Assumptions Taken for Cost Calculation", 2, ORANGE)
    _add_bullets_from_text(doc, tco_assumptions)
    _section_divider(doc)

    # ── 8. Security Best Practices ────────────────────────────────────────
    _heading(doc, "8. AWS Security Best Practices", 1, DARK)
    sec_links = [
        "Best Practices for AWS root users: https://docs.aws.amazon.com/accounts/latest/reference/best-practices-root-user.html",
        "Best Practices for AWS Access Keys: https://docs.aws.amazon.com/accounts/latest/reference/credentials-access-keys-best-practices.html",
        "Shared Responsibility Model: https://aws.amazon.com/compliance/shared-responsibility-model/",
        "AWS CloudTrail: https://aws.amazon.com/cloudtrail/",
        "Trusted Advisor: https://aws.amazon.com/premiumsupport/trustedadvisor/",
        "Enable MFA: https://docs.aws.amazon.com/IAM/latest/UserGuide/id_credentials_mfa.html",
    ]
    for b in sec_links:
        _bullet(doc, b)
    _section_divider(doc)

    _heading(doc, "8.1 AWS Identity and Access Management (IAM)", 2, ORANGE)
    for b in IAM_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    _heading(doc, "8.2 AWS IAM Access Analyzer", 2, ORANGE)
    for b in IAM_ANALYZER_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    _heading(doc, "8.3 Detective Controls", 2, ORANGE)
    for b in DETECTIVE_CONTROLS_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    _heading(doc, "8.4 AWS Detective", 2, ORANGE)
    for b in AWS_DETECTIVE_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    _heading(doc, "8.5 AWS Security Hub", 2, ORANGE)
    for b in SECURITY_HUB_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    # ── 9. Partner Deliverables ───────────────────────────────────────────
    _heading(doc, "9. AWS Partner Deliverables", 1, DARK)
    for b in PARTNER_DELIVERABLES_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    # ── 10. Customer Dependencies ─────────────────────────────────────────
    _heading(doc, "10. Customer Dependencies", 1, DARK)
    for b in CUSTOMER_DEPS_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    # ── 11. Assumptions ───────────────────────────────────────────────────
    _heading(doc, "11. Assumptions", 1, DARK)
    for b in ASSUMPTIONS_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    # ── 12. Exclusions ────────────────────────────────────────────────────
    _heading(doc, "12. Exclusions", 1, DARK)
    for b in EXCLUSIONS_BULLETS:
        _bullet(doc, b)
    _section_divider(doc)

    # ── 13. Risk Analysis ─────────────────────────────────────────────────
    _heading(doc, "13. Risk Analysis", 1, DARK)
    risk_table = doc.add_table(rows=1 + len(RISK_TABLE_DATA), cols=3)
    risk_table.style = "Table Grid"
    for i, h in enumerate(["S. No.", "Risk", "Proposed Mitigation"]):
        cell = risk_table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        _set_cell_bg(cell, "1F2763")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for idx, (sno, risk, mit) in enumerate(RISK_TABLE_DATA):
        row = risk_table.rows[idx + 1]
        row.cells[0].text = sno
        row.cells[1].text = risk
        row.cells[2].text = mit
    _section_divider(doc)

    # ── 14. Project Plan ──────────────────────────────────────────────────
    _heading(doc, "14. Project Plan", 1, DARK)
    pp_table = doc.add_table(rows=1 + len(PROJECT_PLAN_TASKS), cols=6)
    pp_table.style = "Table Grid"
    for i, h in enumerate(["S.No", "Activities / Milestones", "Month 01", "Month 02", "Month 03", "Month 04"]):
        cell = pp_table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        _set_cell_bg(cell, "1F2763")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for idx, (sno, activity) in enumerate(PROJECT_PLAN_TASKS):
        row = pp_table.rows[idx + 1]
        row.cells[0].text = sno
        row.cells[1].text = activity
        if not sno:  # section header row
            _set_cell_bg(row.cells[1], "F2F2F2")
            if row.cells[1].paragraphs[0].runs:
                row.cells[1].paragraphs[0].runs[0].bold = True
    _section_divider(doc)

    # ── 15. Commercial Terms ──────────────────────────────────────────────
    _heading(doc, "15. Commercial Terms and Conditions", 1, DARK)
    for line in COMMERCIAL_TERMS.split("\n"):
        p = doc.add_paragraph()
        p.add_run(line).font.size = Pt(10)

    doc.add_paragraph()

    # Payment table
    pay_table = doc.add_table(rows=6, cols=2)
    pay_table.style = "Table Grid"
    pay_data = [
        ("Beneficiary Name", "OPERISOFT TECHNOLOGIES PVT LTD"),
        ("Bank Name & Address", "Yes BANK, Vashi Branch, Navi Mumbai"),
        ("Bank A/c No.", "073838389703"),
        ("Account Type", "Current Account"),
        ("IFSC/IMPS/RTGS CODE", "IC7830151"),
        ("SWIFT CODE", "I7j9fifINBBCTS"),
    ]
    for i, (k, v) in enumerate(pay_data):
        pay_table.rows[i].cells[0].text = k
        pay_table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
        pay_table.rows[i].cells[1].text = v

    doc.add_paragraph()
    _para(doc, "Our (GST Number) GSTIN # 27A78787R45M1ZZ", bold=True)
    doc.add_paragraph()

    # Disclaimer
    _heading(doc, "DISCLAIMER:", 2, DARK)
    disc_bullets = [
        "Training or anything else which is not mentioned in above document is out of scope.",
        "For any feature/functionalities to be added which is not in scope will be charged extra as per mutual agreement.",
        "Using the documents or data for any purpose other than this, should be taken in writing by Operisoft.",
    ]
    for b in disc_bullets:
        _bullet(doc, b)

    doc.add_paragraph()
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run("Operisoft Technologies Pvt Ltd")
    run.bold = True
    run.font.color.rgb = DARK

    footer_p2 = doc.add_paragraph()
    footer_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_p2.add_run("CIN: U72900MH2022PTC384302 | GSTIN: 27AADCO8045M1ZZ | www.operisoft.com | E: info@operisoft.com").font.size = Pt(8)

    # ── Save file ──────────────────────────────────────────────────────────
    output_dir = tempfile.gettempdir()
    filename = f"SOW_{customer_name.replace(' ', '_')}.docx"
    output_path = os.path.join(output_dir, filename)
    doc.save(output_path)
    return output_path
