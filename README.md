# gs-jira
Automation script between Google Sheet and JIRA


## Configuring and Running Script


##### 1. First, Create virtualenv

```bash
virtualenv venv
```

##### 2. Activate virtualenv

```bash
source venv/bin/activate
```

##### 3. Install python packages

```bash
pip install -r requirements.txt
```

##### 4. Create .env from .env.example and fill in proper values

##### 5. Run script
```bash
python gs2jira.py
```


## Python Google sheet API

- https://github.com/burnash/gspread
- https://gspread.readthedocs.io/en/latest/oauth2.html

## Python JIRA library
- https://jira.readthedocs.io/en/master/index.html
