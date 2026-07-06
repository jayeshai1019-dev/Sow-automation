from pricing_client import build_pricing_estimate

MOM = """
Client wants to migrate their existing on-premise data warehouse to AWS.
Existing databases include SQL Server (6 TB) and Oracle (3 TB).
Monthly CSV files (~500 GB) are received from external vendors via SFTP.
ETL processing should be performed using AWS Glue.
Processed data should be stored in Amazon Redshift.
Raw data should be stored in Amazon S3.
Security requirements include KMS encryption, Secrets Manager, CloudTrail, and CloudWatch.
Data should be retained for 7 years. Disaster Recovery should be considered.
"""

result = build_pricing_estimate("Data Warehouse Client", MOM)
print("\n=== RESULT ===")
print("URL:    ", result["url"])
print("Error:  ", result["error"])
print("Services in SOW table:", len(result["services"]))
for s in result["services"]:
    print(f"  {s['name']}: {s['prod']['spec']}")
