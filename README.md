# q-sharepoint-api
API til sharepoint

# 🔐 SharePoint Authentication – APP vs USER

Denne løsning bruger **to forskellige login‑metoder** til SharePoint:

- ✅ APP login (Azure App / client_credentials)
- ✅ USER login (teknisk bruger / password)

Det er vigtigt at forstå, hvornår man bruger hvad.

---

# 🧠 Overblik

| Funktion | Login |
|--------|------|
| Læse liste items (GET) | ✅ APP |
| Oprette / opdatere items | ✅ APP |
| Excel (Graph) | ✅ APP |
| Attachments (liste vedhæftninger) | ✅ USER |

---

# ✅ APP LOGIN (client_credentials)

## 🔹 Hvad er det?

APP login bruger Azure App:

```python
grant_type = "client_credentials"

