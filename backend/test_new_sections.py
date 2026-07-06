"""
Test script to verify the new High Level Scope and Scope of Work sections
"""
import os
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

from bedrock_client import generate_high_level_scope, generate_scope_of_work

# Sample test data
customer_name = "ABC Corporation"
mom_text = """
- Customer wants to implement a GenAI-based PDF Data Extraction and Automation Solution
- Web-based UI for uploading PDF documents
- AWS Amplify for hosting frontend
- Amazon S3 for document storage
- AWS Lambda for PDF processing
- Amazon Bedrock for AI-driven data extraction
- Excel output generation for users
- This is a POC project
- Need to validate accuracy during POC phase
"""

print("="*80)
print("Testing High Level Scope Generation (3.1)")
print("="*80)
high_level_scope = generate_high_level_scope(customer_name, mom_text)
print(high_level_scope)
print()

print("="*80)
print("Testing Scope of Work Generation (3.2)")
print("="*80)
scope_of_work = generate_scope_of_work(customer_name, mom_text)
print(scope_of_work)
print()

print("="*80)
print("Test completed successfully!")
print("="*80)
