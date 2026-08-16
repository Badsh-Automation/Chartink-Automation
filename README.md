# 📊 Chartink Dashboard Automation (GitHub Actions + Google Drive)

## 🎯 Kya Hai Yeh?
Roz **5:00 PM IST** pe automatic:
1. Chartink se files download
2. Power Query jaisa processing
3. 8 sheets wala Dashboard generate
4. **Google Drive** pe upload

---

## 📁 Files Structure (Repo Mein)

```
chartink-dashboard/           <-- Aapka GitHub Repo
├── .github/
│   └── workflows/
│       └── daily-chartink.yml   <-- Schedule file
├── chartink_complete.py         <-- Main script
├── requirements.txt             <-- Libraries
└── README.md                    <-- Yeh file
```

---

## 🚀 Step-by-Step Setup

### STEP 1: GitHub Repo Banana

1. [github.com](https://github.com) pe jao → Sign up/Login
2. **New Repository** → Name: `chartink-dashboard`
3. **Private** select karo → **Create repository**

### STEP 2: Files Upload Karna

Repo mein yeh 3 files upload karo:
- `chartink_complete.py`
- `requirements.txt`
- `.github/workflows/daily-chartink.yml`

**Kaise?**
- Repo page pe **"Add file"** → **"Upload files"**
- Ya **drag & drop** karo

### STEP 3: GitHub Secrets Add Karna (Very Important!)

Repo page pe:
```
Settings → Secrets and variables → Actions → New repository secret
```

Yeh **4 secrets** add karo:

| Secret Name | Value | Kya Hai |
|-------------|-------|---------|
| `CHARTINK_EMAIL` | aapka@email.com | Chartink login email |
| `CHARTINK_PASSWORD` | aapka_password | Chartink password |
| `GOOGLE_DRIVE_CREDENTIALS` | {poora JSON} | Service account JSON (Step 4 mein milega) |
| `GOOGLE_DRIVE_FOLDER_ID` | 1AbC...xYz | Google Drive folder ID (Step 5 mein milega) |

---

### STEP 4: Google Drive Setup (Service Account)

#### 4.1 Google Cloud Project Banana
1. [console.cloud.google.com](https://console.cloud.google.com) pe jao
2. **Select a project** → **New Project**
3. Name: `ChartinkDashboard` → **Create**

#### 4.2 Google Drive API Enable Karna
1. Left menu → **APIs & Services** → **Library**
2. Search: **Google Drive API**
3. Click → **Enable**

#### 4.3 Service Account Banana
1. Left menu → **APIs & Services** → **Credentials**
2. **Create Credentials** → **Service Account**
3. Name: `chartink-uploader` → **Create and Continue**
4. Role: **Basic** → **Editor** → **Continue**
5. **Done**

#### 4.4 JSON Key Download Karna
1. Credentials page pe service account pe click karo
2. **Keys** tab → **Add Key** → **Create new key**
3. **JSON** select → **Create**
4. JSON file download hogi → **SAVE KARO**

#### 4.5 GitHub Secret Mein Daalna
JSON file kholiye, **poora content copy** karo:
```json
{
  "type": "service_account",
  "project_id": "...",
  ...
}
```

GitHub secret `GOOGLE_DRIVE_CREDENTIALS` mein **poora JSON paste** karo.

---

### STEP 5: Google Drive Folder Banana

#### 5.1 Folder Create Karna
1. [drive.google.com](https://drive.google.com) pe jao
2. **New** → **Folder** → Name: `Chartink Dashboards`
3. Folder kholiye

#### 5.2 Folder ID Nikalna
URL mein dekhiye:
```
https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz
```
`1AbCdEfGhIjKlMnOpQrStUvWxYz` yeh hai **Folder ID**

#### 5.3 Service Account Ko Share Karna
1. Folder pe **right-click** → **Share**
2. Email add karo: `chartink-uploader@chartinkdashboard.iam.gserviceaccount.com`
   (yeh aapke service account ka email hai)
3. **Editor** role do → **Send**

#### 5.4 GitHub Secret Mein Daalna
Folder ID ko `GOOGLE_DRIVE_FOLDER_ID` secret mein daal do.

---

### STEP 6: Test Karna (Manual Run)

Repo page pe:
```
Actions tab → Daily Chartink Dashboard → Run workflow
```

**Run workflow** button dabao → **Run workflow**

5-10 minute wait karo, phir:
- **Green tick** = Success ✅
- **Red cross** = Error ❌ (logs check karo)

---

### STEP 7: Daily Schedule Verify Karna

Schedule already set hai:
```yaml
cron: '30 11 * * *'  # Roz 5 PM IST
```

**Time verify karo:**
- GitHub UTC time mein chalta hai
- 11:30 AM UTC = 5:00 PM IST (India time)
- Agar DST change ho toh 4:30 PM ho sakta hai

---

## 📧 Agar Koi Problem Aaye

### Problem 1: Chrome/Driver Error
```
Logs mein: "Chrome not found" ya "driver error"
```
**Solution:** GitHub Actions Ubuntu image mein Chrome pre-installed hota hai. Script mein `chromium-browser` use ho raha hai.

### Problem 2: Chartink Login Fail
```
Logs mein: "Login failed"
```
**Solution:**
- Secrets sahi hain verify karo
- Chartink account active hai?
- 2FA enabled hai toh disable karo ya app password use karo

### Problem 3: Google Drive Upload Fail
```
Logs mein: "Google Drive upload failed"
```
**Solution:**
- `GOOGLE_DRIVE_CREDENTIALS` mein poora JSON hai?
- Service account ko folder share kiya hai?
- Drive API enabled hai?

### Problem 4: Schedule Nahi Chal Raha
```
Actions mein scheduled run nahi dikh raha
```
**Solution:**
- Pehli baar manual run karna padta hai
- GitHub free mein 60 days inactive repo pe schedule band kar deta hai
- Roz manual run ya commit karte raho

---

## 📊 Output Kya Milega?

Google Drive `Chartink Dashboards` folder mein:
```
Dashboard_2026-08-15.xlsx
Dashboard_2026-08-16.xlsx
Dashboard_2026-08-17.xlsx
...
```

**Roz nayi file** ya **same file update** hoti hai (configurable).

---

## ✅ Final Checklist

- [ ] GitHub repo created (private)
- [ ] 3 files uploaded
- [ ] 4 secrets added
- [ ] Google Cloud project created
- [ ] Drive API enabled
- [ ] Service account created
- [ ] JSON key downloaded & added to secrets
- [ ] Drive folder created & shared with service account
- [ ] Folder ID added to secrets
- [ ] Manual test run successful
- [ ] Google Drive pe file aayi

---

**Bas ho gaya! Roz 5 PM pe auto chalega!** 🚀
