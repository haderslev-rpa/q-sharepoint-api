import requests


class CopilotAPI:

    def __init__(self, token):
        self.base = "https://graph.microsoft.com/beta/copilot"
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

    def start_conversation(self):
        url = f"{self.base}/conversations"

        r = requests.post(url, headers=self.headers, json={})
        r.raise_for_status()

        return r.json()

    def send_message(self, conversation_id, body):
        url = f"{self.base}/conversations/{conversation_id}/chat"

        r = requests.post(url, headers=self.headers, json=body)
        r.raise_for_status()

        return r.json()