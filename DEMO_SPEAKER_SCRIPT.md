# UC-01 — Demo Speaker Script (~8 minutes)

**Framing rule for the whole talk:** this is an *early-stage PoC / workpaper assistant*, not a finished product and not an RCM tool. Under-promise. Every number below is offered as an early estimate "to be validated," and you say so out loud at least once.

Legend: **[DO]** = on-screen action · *(aside)* = optional ad-lib · timings are cumulative.

---

## 0:00 – 0:45 · Open (the hook)

> "Quick show of hands — who here *enjoys* writing the CDE and COE sections of a control workpaper? … Right, nobody, and the one person who raised their hand, we need to talk.
>
> So this is a small side-of-desk PoC I've been calling UC-01 — a **workpaper assistant**. To be clear up front: it does **not** replace the auditor's judgment, and it's **early-stage** — think 'promising intern', not 'senior manager'. What it does is take the boring 80% — turning a walkthrough into a structured draft workpaper — so we spend our time on the 20% that actually needs an auditor's brain."

*(If the earlier RCM tool was just demoed:)*
> "And this pairs nicely with the RCM tool you just saw — that one's brilliant at the RCM; mine picks up *after* the RCM and drafts the rest of the workpaper: CDE, COE, DA, exceptions, conclusion. Complementary, not competing."

---

## 0:45 – 1:45 · What it is + inputs

> "The idea: an audit has controls; each control has a walkthrough. These days our walkthroughs have **two** artefacts — the Teams transcript *and* the Copilot meeting summary, because Copilot's on during the call. The summary is honestly gold: it already lists attendees, key topics and the RFIs we raised."

**[DO]** Open **`Copilot_Summary_DB_Access_Mgmt_Walkthrough.docx`**. Point at Attendees + Action items/RFIs.

> "So the assistant ingests the Copilot summary *and* the transcript. From the transcript it captures the stuff that has to be in the workpaper — date, attendees, who said what."

**[DO]** Open **`Transcript_DB_Access_Mgmt_Walkthrough.txt`** for a second — show the attendee header and one line (e.g. the MongoDB discovery gap admission).

> "This is a real control area — Database *Access* Management: CyberArk vaulting, Imperva DAM, recertification, across 342 production databases."

---

## 1:45 – 4:30 · Live demo

**[DO]** In the app sidebar → **🎬 Demo data → Load Database Access Management demo**. Select the control **Database Privileged Access Management**.

> "Everything I'm about to click runs **locally and deterministically** — no external calls, no internet, nothing leaves the laptop. I'll explain why that matters in a minute."

**Documents tab**
> "Walkthrough Copilot summary and transcript are already attached. I click **Build workpaper draft**…"

**[DO]** Click **Build RCM from documents**. Switch to **RCM tab**.
> "…and it drafts the control matrix row — the control described in our WHO-WHEN-WHAT-HOW-WHY style, the key questions, the CDE/COE/DA procedures. I'm *not* dwelling here — the RCM's not my headline. Note it's fully editable; the auditor always has the pen."

**Testing Phases tab** — the star of the show.
**[DO]** CDE → **Run CDE analysis**.
> "CDE — control design. Notice it's captured the **walkthrough date and attendees** right in the workpaper, then design strengths and — the useful part — design **gaps**: MongoDB sits outside CyberArk discovery, no same-team segregation-of-duties check, no periodic recertification for DBA accounts."

**[DO]** COE → **Run COE analysis**.
> "COE — operating effectiveness, GIA sampling. Sample of 25. It's pulled the exceptions: the self-approval SoD pattern — 3 of 18 — one late revocation, and the recertification gap."

**[DO]** Open **`COE_Testing_Results_DB_Access_Mgmt.xlsx`** (the evidence file) beside it.
> "And here's the actual COE testing workbook it lines up with — every test step, pass/fail, exceptions, evidence references, conclusion. Same for CDE." *(briefly show `CDE_Testing_Results…xlsx` and `TIP_Database_Access_Management.pdf`)*

**[DO]** DA → **Run DA analysis**.
> "Data analytics — the vaulting reconciliation: 1,247 privileged accounts, 49 unvaulted, 3.9%, and it spots that the MongoDB slice is a *new* finding management's own 2.4% number missed."

**[DO]** EXCEPTIONS → **Run** → then **Conclusion → Generate** → **Export**.
> "It consolidates the exceptions into a log with ratings and recommended actions, drafts an overall conclusion — Partially Effective, two High-rated issues — and exports a **Word workpaper** and an **Excel RCM**. Auditor reviews, edits, done."

---

## 4:30 – 5:30 · Under the hood (be honest here)

> "How does it work *today*? Two honest points.
>
> One — **right now the drafting is rule-based, not a live AI model.** I'll say that plainly because it's a feature, not a bug: it means the PoC runs entirely inside the bank, deterministically, with zero data leaving — which is exactly what let me demo it to you today without a six-week approval.
>
> Two — it's **grounded in historical audit data**."

**[DO]** Sidebar → **📚 Knowledge base**. Show the counts.
> "Past control workpapers, past issues, our methodology library — that's what makes the drafts sound like *our* workpapers and not generic ChatGPT. And the human-in-the-loop is hard-wired: nothing lands in the workpaper without the auditor accepting it."

---

## 5:30 – 6:30 · Honest metrics + how close to the real thing

> "Numbers — and I want to be careful not to over-sell these; they're early estimates to validate:
>
> - A control workpaper first-draft is roughly **2–3 hours** of manual writing. The assistant gets you a reviewable draft in **minutes** — call it a **~60–70% cut in first-draft time**, to be measured properly on real audits.
> - About **~75% of the draft is auto-populated** — structure, narrative, exceptions scaffold — leaving the auditor the judgment 25%.
> - Compared with the *actual* signed workpaper, this draft is roughly **80–85% there** on structure and content.
>
> The missing 15–20% is the honest part. Key differences:
> 1. It doesn't yet *do the testing* — it drafts around evidence the auditor still gathers and concludes on.
> 2. It doesn't yet parse structured evidence (the CDE/COE spreadsheets) automatically.
> 3. Formatting isn't SARA-native yet.
> 4. It doesn't auto-link to issues/RFIs.
>
> Every one of those closes with the **live model** and proper integration — which is the roadmap."

---

## 6:30 – 7:30 · Roadmap (the ask)

> "Two next steps.
>
> **First — AI Factory.** To swap the rule-based engine for a real model — Claude, GPT, Qwen — the use case and PoC code go to **AI Watchtower** for review, then the AI Factory team reviews and sets up a session with us. Realistically a **1–2 month** process, and their developers take it from PoC to production. The app's already built for this: there's an **API-key switch** — the day we're onboarded, I paste the Factory key and the same buttons run real analysis, no rework."

**[DO]** *(optional)* open the sidebar **🔑 Live analysis** panel for one second.

> "**Second — SARA.** Long-term this shouldn't be a standalone app; it should live inside SARA, where audits, issues and reports already are. I checked with the SARA technical team — their stack is **Lit front-end, Quarkus back-end, Postgres**. Handy coincidence: I've already prototyped this exact assistant on **that identical stack** — Lit, Quarkus, Postgres — so integration is largely lifting the module into SARA as a panel on the audit record, reading and writing to SARA's Postgres, with the AI calls routed through AI Factory. Not a rebuild — a port."

---

## 7:30 – 8:00 · Close

> "So — early PoC, deliberately runs safe and offline today, drafts ~75% of a control workpaper grounded in our own history, and there's a clean path: AI Factory for the brain, SARA for the home.
>
> What I'd love from this room: a nod to submit it to AI Watchtower, and a pilot control or two on a live audit to measure the real time-saving.
>
> And if it saves each of us even one Sunday-evening workpaper session a quarter… I'll consider the intern promoted."

---

# Q&A — anticipated questions & crisp answers

**"Is our data going to an external AI?"**
> Not in this PoC — it's fully local and rule-based, nothing leaves the environment. When we move to a live model it's exclusively through **AI Factory**, in-tenant and governed; that's the whole point of the Watchtower review.

**"How is this different from the RCM tool [colleague] built?"**
> Complementary. Theirs focuses on the RCM; mine drafts the **rest of the workpaper** — CDE, COE, DA, exceptions, conclusion — and exports it. Ideally they feed each other.

**"Isn't it just ChatGPT with extra steps?"**
> Two differences: it's grounded in **our** historical workpapers and methodology so it speaks GIA, and it's **deterministic and offline** today. A generic chatbot is neither auditable nor bank-safe.

**"What if the AI hallucinates a finding?"**
> Human-in-the-loop is mandatory — nothing enters the workpaper without the auditor accepting it. Today's output is curated from real historical patterns, so there's no fabrication; with the live model, the auditor remains the reviewer and every draft cites the evidence it drew on.

**"How accurate is it really?"**
> On this control the draft is ~80–85% of the final workpaper's structure and content. It won't *conclude* for you — it drafts, you test and sign.

**"How long to production?"**
> Gated by AI Factory — ~1–2 months for onboarding/review, then their developers productionise. The SARA integration is a follow-on, made easier because I've already built it on SARA's exact stack.

**"What did *you* build vs. the AI?"**
> I built the app, the workflow, the historical-data grounding and the exports. The model (once onboarded) does the drafting. Today that drafting is rule-based logic I wrote from past workpapers.

**"Does it handle other audits, or just databases?"**
> The engine is control-agnostic; I've templated it on Database Access Management because I had the richest historical data there. New domains need their historical workpapers loaded — that's the grounding step, not a rebuild.

**"What does it cost?"**
> Effectively zero today (runs locally). Live-model cost is AI Factory's consumption model — a few cents to a couple of dollars per workpaper draft, to confirm during onboarding.

**"Can I break it live?"** *(if someone clicks around)*
> Please do — it's a PoC. Worst case I reload the demo data. *(Keep it light.)*

---

### Pre-flight checklist (do this 10 min before)
- [ ] `streamlit run app.py` — sidebar shows **🔵 Demo mode**.
- [ ] Loaded the demo audit once; ran CDE on the control so you know the click works.
- [ ] Six test files open in tabs/Finder ready to show: Copilot summary, transcript, CDE xlsx, COE xlsx, TIP pdf, access matrix.
- [ ] Delete `uc01.db` for a clean slate if you re-seed.
