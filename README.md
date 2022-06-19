
# GMAIL Parser to JSON

Python project to get Gmail body, from, to, subject, and also parse CSV attachment to payload.json file locally.
The file credentials.json must be downloaded from https://console.cloud.google.com/apis/credentials


A project must be created as explained in Google Cloud console and it must be shifted from testing to published.
https://developers.google.com/gmail/api/quickstart/python

## Usage/Examples

Usage:
Install all the dependencies needed


```
pip install pipreqs

pip install -r requirements.txt
```

Alternatively you can install the following:
```
pip install google_api_python_client
pip install google_auth_oauthlib
pip install pandas
pip install protobuf
pip install xlrd
pip install BeautifulSoup
```

On first run, authentication is done in a browser window to the required GMAIL account and a file token.json is created. On second run the program runs successfully to create payload.json with the email data.

