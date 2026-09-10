#!/usr/bin/env python
# coding: utf-8

# In[ ]:





# In[1]:


import requests


def get_access_token():
    """Gets an access token for the Microsoft Graph API."""

    url = "https://login.microsoftonline.com/common/oauth2/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": "client_key",
        "client_secret":"client_secret",
        "resource": "https://graph.microsoft.com/"
    }
    response = requests.post(url, data=data)
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        # print(response.status_code)
        return None


token = get_access_token()
r = requests.get("https://graph.microsoft.com/v1.0/me/messages",
                 headers={"Authorization": "Bearer " + token})

if r.status_code == 200:
    for message in r.json():
        print(message["subject"])
else:
    print("Authentication Failed")


# In[ ]:




