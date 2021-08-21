# pip install boto3
# pip install botocore
import os
import json
import boto3

from boto3.session import Session
from botocore.exceptions import ClientError

REGION = os.environ.get('AWS_DEFAULT_REGION', 'sa-east-1')
BUCKET_NAME = os.environ.get('BUCKET_NAME', 'blackdevs-aws')

session = Session(region_name=REGION)
s3 = session.resource('s3')

# download do arquivo
try:
  s3.Bucket(BUCKET_NAME).download_file('s3-file-python-vba/file.json', 'C:\\Users\\julio\\Documents\\Excel-PowerBI\\s3-file-python-vba\\file.json')

except ClientError as e:
  if e.response['Error']['Code'] == "404":
    print("O arquivo não existe")
  else:
    raise

# lê o arquivo
with open(f"./file.json", "r") as file:
  text = file.read()
  print(json.loads(text))
  file.close()
