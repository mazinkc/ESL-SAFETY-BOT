# ESL Safety Observation Bot — Setup Guide
## Zero Cost | Firebase + Telegram | Render Hosting

---

## STEP 1 — Create Telegram Bot (5 minutes)

1. Open Telegram → search **@BotFather**
2. Send `/newbot`
3. Name it: `ESL Safety Bot`
4. Username: `esl_safety_obs_bot` (must be unique, try variations)
5. BotFather gives you a **token** like:
   `7123456789:AAFxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx`
6. Copy and save this — this is your `BOT_TOKEN`

---

## STEP 2 — Create Firebase Project (10 minutes)

1. Go to **https://console.firebase.google.com**
2. Click **Add Project** → name it `esl-safety-bot`
3. Disable Google Analytics (not needed) → Create

### Enable Firestore
4. Left sidebar → **Firestore Database** → Create Database
5. Select **Start in test mode** → choose region `asia-south1` (Mumbai) → Enable

### Enable Storage
6. Left sidebar → **Storage** → Get Started
7. Start in **test mode** → choose same region → Done

### Get Service Account Key
8. Go to **Project Settings** (gear icon top left)
9. Click **Service Accounts** tab
10. Click **Generate New Private Key** → Download JSON file
11. Open that JSON file → copy ALL its contents
    This is your `FIREBASE_CREDS_JSON`

### Get Storage Bucket Name
12. Go to **Storage** → copy the bucket name shown at top
    Looks like: `esl-safety-bot.appspot.com`
    This is your `FIREBASE_BUCKET`

---

## STEP 3 — Deploy on Render (10 minutes)

1. Go to **https://render.com** → Sign up free with GitHub

2. Push your bot files to a GitHub repo:
   ```
   esl_safety_bot/
   ├── main.py
   ├── report_excel.py
   ├── report_pptx.py
   ├── requirements.txt
   └── SETUP.md
   ```

3. In Render → **New** → **Web Service**
4. Connect your GitHub repo
5. Settings:
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python main.py`

6. Add Environment Variables (click **Add Environment Variable**):

   | Key | Value |
   |-----|-------|
   | `BOT_TOKEN` | Your Telegram bot token from BotFather |
   | `FIREBASE_CREDS_JSON` | Paste the entire service account JSON |

   > ✅ No need to set FIREBASE_BUCKET — hardcoded as apex-esl.firebasestorage.app

7. Click **Create Web Service** → Wait 2-3 minutes for deploy

---

## STEP 4 — Test It

1. Open Telegram → search your bot name
2. Send `/start` — should show help message
3. Forward a test observation:

```
Name: Test Officer
Date/Time: 09/05/2026 10:30
Observation: Test loose wire near conveyor
Area: Sinter Plant
Responsible: Suresh Singh
Status: Open
Target Date: 15/05/2026
```

4. Bot should reply: ✅ Observation Saved! with Ref No
5. Send `/stats` → should show count
6. Send `/report` → should receive Excel + PPTX files

---

## STEP 5 — Pin Format in WhatsApp Group

Pin this message in your WhatsApp group:

```
📋 SAFETY OBSERVATION FORMAT
━━━━━━━━━━━━━━━━━━━━━
Name:
Date/Time:
Observation:
Area:
Responsible:
Status: (Open / In Progress / Closed)
Target Date:
━━━━━━━━━━━━━━━━━━━━━
Attach photo with message
```

---

## DAILY WORKFLOW

1. Officers post in WhatsApp group (as usual)
2. You forward each observation to the Telegram bot (10 seconds each)
3. Bot auto-saves to Firebase
4. End of week → type `/report` → receive PPTX + Excel instantly

---

## COMMANDS REFERENCE

| Command | What it does |
|---------|-------------|
| `/start` | Show help + format |
| `/stats` | Live count (Total / Open / Closed) |
| `/report` | Generate and send PPTX + Excel |

---

## FREE TIER LIMITS (You'll Never Hit These)

| Resource | Free Limit | Your Usage |
|----------|-----------|-----------|
| Firestore reads | 50,000/day | ~100/day |
| Firestore writes | 20,000/day | ~20/day |
| Firebase Storage | 5 GB | Months of photos |
| Render hosting | 750 hrs/month | ~720 hrs |
| Telegram Bot API | Unlimited | — |

**Total Monthly Cost: ₹0**
