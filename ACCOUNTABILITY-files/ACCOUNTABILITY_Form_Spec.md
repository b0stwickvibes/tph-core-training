# ACCOUNTABILITY FORM — SPECIFICATION
**Three Points Hospitality — C.O.R.E. Training System**
**Last Updated: March 16, 2026**

---

## PURPOSE

The Accountability Form is submitted by the **trainer** at the end of every training shift. It is the digital paper trail that runs alongside the signed physical checklist. Together, they confirm the shift happened, the trainer covered the material, and the trainee was assessed.

This is a **trainer accountability tool** — not a trainee self-assessment.

---

## FORM OWNERSHIP

| Field | Detail |
|---|---|
| **Submitted by** | Trainer — end of every training shift, before leaving the building |
| **Reviewed by** | Devin (weekly), GM (weekly audit) |
| **Enforcement** | No form + no checklist photo = no trainer incentive for that shift |
| **Storage** | Google Sheets backend; GM and Devin have full access |

---

## SECTION 1: SESSION INFO (auto + dropdowns)

| Field | Type | Notes |
|---|---|---|
| Location | Dropdown | Cantina Añejo, Original American Kitchen, White Buffalo |
| Trainer Name | Dropdown | Filtered by location (pulls from Trainers config) |
| Trainee Name | Text input | Free text |
| Position | Dropdown | Bartender, Server, Host |
| Training Day | Dropdown | Day 1 through Day 5 |
| Shift | Dropdown | Morning, Afternoon, Evening (auto-selected by time of day) |
| Timestamp | Auto | Captured on submission |

---

## SECTION 2: TRAINEE PERFORMANCE SCORES

Five categories, each scored 1–5. These are the trainer's assessment of the trainee — not self-reported.

| Category | What It Measures |
|---|---|
| **Knowledge & Understanding** | Can the trainee explain what they learned today? Do they understand the *why*, not just the *what*? |
| **Technical Skills & Execution** | Did the trainee physically execute the skills covered? Pour accuracy, POS usage, table management — whatever applies to the day. |
| **Customer Service & Communication** | How did the trainee interact with guests? Professional, warm, confident — or awkward, quiet, avoidant? |
| **Teamwork & Collaboration** | Did the trainee communicate with the team? Help without being asked? Or operate in a silo? |
| **Professionalism & Attitude** | Punctuality, phone usage, attitude toward correction, willingness to learn. Body language and effort level. |

**Scoring:**
- **1** = Poor — did not demonstrate this at all
- **2** = Below Average — inconsistent, significant gaps
- **3** = Average — met baseline expectations for this training day
- **4** = Good — exceeded expectations in some areas
- **5** = Excellent — performed at or near solo-shift readiness for this category

**Auto-calculated:**
- Total Score: ___ / 25
- Percentage: ___%
- Performance Level: Excellent (≥90%), Good (75–89%), Needs Improvement (<75%)

---

## SECTION 3: TRAINER ACCOUNTABILITY QUESTIONS

These three questions are required on every submission. They map directly to the program's accountability structure.

### Question 1: Coverage Confirmation
> **"What did you cover today? List the main topics from the checklist."**

- Field type: Open text (required, minimum 30 characters)
- Purpose: Forces the trainer to confirm what was actually taught — not just check boxes
- Red flag: Answer that doesn't map to the day's checklist topics

### Question 2: Trainee Gaps
> **"Where is the trainee struggling? What needs extra attention?"**

- Field type: Open text (required, minimum 20 characters)
- Purpose: Surfaces gaps early so they can be addressed before Day 5
- Red flag: "Nothing" or vague answers on Day 1–3 followed by a Needs Improvement score on Day 5

### Question 3: Plan Forward
> **"What's the plan for the next training shift?"**

- Field type: Open text (required, minimum 20 characters)
- Purpose: Confirms the trainer is thinking ahead, not winging it
- Red flag: Generic answers that don't reference specific gaps from Q2

---

## SECTION 4: RECAP CONFIRMATION

> **"Did you complete the End-of-Shift Recap with the trainee before they clocked out?"**

- Field type: Yes / No toggle
- If **NO**: follow-up field appears — "Why not?" (required text)
- Purpose: The end-of-shift recap is a non-negotiable line item on every floor checklist. This confirms it happened.
- Flag: Any NO response triggers a review. Repeated NOs = trainer accountability conversation.

---

## SECTION 5: CHECKLIST PHOTO UPLOAD

> **"Upload a photo of the signed Day [X] checklist."**

- Required field — form cannot be submitted without a photo
- Accepted formats: JPG, PNG, HEIC
- File size limit: 10MB
- Photo is stored in Google Drive (organized by location/trainee/day)
- Reviewed in GM weekly audit — must be legible and signed by both trainer and trainee

---

## SECTION 6: FORM FLOW

| Step | What Happens |
|---|---|
| 1 | Trainer opens form (web app via Google Sheets deployment) |
| 2 | Select Location → Trainer Name auto-filters |
| 3 | Enter Trainee Name, Position, Training Day, Shift |
| 4 | Rate trainee on 5 categories (1–5 each) |
| 5 | Score summary auto-displays (total, percentage, performance level) |
| 6 | Answer 3 accountability questions (text fields) |
| 7 | Confirm End-of-Shift Recap (Yes/No) |
| 8 | Upload checklist photo (required) |
| 9 | Submit |
| 10 | Confirmation screen: "Assessment submitted. Record ID: TR-XXXXXXXX" |

---

## SECTION 7: FLAG LOGIC (alerts for Devin / GM)

| Trigger | What Happens |
|---|---|
| **Day 4 or Day 5 score < 75%** | Email alert to Devin. Trainee may not be ready for mock service. Check in before clearing. |
| **Recap = NO** | Flagged in GM weekly audit. Repeated NOs = trainer accountability conversation. |
| **No submission filed for a scheduled training shift** | GM follows up with trainer same day. Shift flagged as incomplete. No incentive for that shift. |
| **Photo missing or illegible** | Trainer resubmits corrected photo within 24 hours. |
| **Consecutive "Needs Improvement" scores (same trainee, 2+ days)** | Devin check-in with trainer. Is this a trainee issue or a training execution issue? |

---

## SECTION 8: INCENTIVE GATE

The accountability form is directly tied to trainer pay:

| Condition | Incentive Status |
|---|---|
| Form submitted + photo uploaded + all fields complete | ✅ Eligible for per-shift incentive |
| Form submitted but photo missing | ❌ Not eligible until photo is resubmitted |
| No form submitted | ❌ Not eligible. Shift treated as incomplete. |
| Trainee completes all 5 days + passes mock service | ✅ Completion bonus triggered |
| Trainee passes Day 30 evaluation | ✅ 30-day performance bonus triggered |

PAID VALIDATION sheet tracks which trainers have been paid per month.

---

## SECTION 9: GM WEEKLY AUDIT PROTOCOL

Every Monday, the GM reviews the prior week's submissions:

1. Pull all submissions from the past 7 days
2. Check for missing submissions (training days with no form filed)
3. Review all flagged entries (low scores, NO recap responses, consecutive struggles)
4. Check photo quality — confirm checklists are signed and legible
5. Follow up with trainer on any missing or flagged entries within 24 hours
6. Log audit completion

> **Target: 100% submission rate.** Any week below 90% triggers a process review.

---

## SECTION 10: DATA ARCHITECTURE

### Training Records Sheet (main data)
Each submission creates one row:

| Column | Source |
|---|---|
| Timestamp | Auto-generated |
| Record ID | Auto-generated (TR-XXXXXXXX) |
| Location | Form: Section 1 |
| Trainer Name | Form: Section 1 |
| Trainee Name | Form: Section 1 |
| Position | Form: Section 1 |
| Training Day | Form: Section 1 |
| Shift | Form: Section 1 |
| Knowledge Score (1–5) | Form: Section 2 |
| Technical Score (1–5) | Form: Section 2 |
| Service Score (1–5) | Form: Section 2 |
| Teamwork Score (1–5) | Form: Section 2 |
| Professionalism Score (1–5) | Form: Section 2 |
| Total Score | Auto-calculated |
| Percentage | Auto-calculated |
| Performance Level | Auto-calculated |
| What Was Covered | Form: Section 3, Q1 |
| Where Struggling | Form: Section 3, Q2 |
| Plan for Next Shift | Form: Section 3, Q3 |
| Recap Completed | Form: Section 4 (Yes/No) |
| Recap Missed Reason | Form: Section 4 (if No) |
| Checklist Photo URL | Form: Section 5 (Drive link) |

### Supporting Sheets (auto-generated)
- **Analytics Dashboard** — Aggregate scoring trends, submission counts, performance distribution
- **Location Summary** — Per-location breakdowns (submissions, avg scores, completion rates)
- **Trainer Performance** — Per-trainer scoring trends, submission frequency, flag counts
- **PAID VALIDATION** — Monthly incentive tracking with payment checkboxes

---

## SECTION 11: IMPLEMENTATION STATUS

### What's Already Built (Live App)
- ✅ Session Info fields (Location, Trainer, Trainee, Position, Day, Shift)
- ✅ 5 scoring categories with 1–5 radio buttons
- ✅ Auto-calculated total, percentage, performance level
- ✅ Overall Notes text field
- ✅ Duplicate submission prevention (MD5 hash)
- ✅ Auto-shift selection by time of day
- ✅ Location-based trainer filtering
- ✅ Training Records sheet with all score columns
- ✅ Analytics Dashboard auto-generation
- ✅ Location Summary sheet
- ✅ Trainer Performance sheet
- ✅ PAID VALIDATION sheet with formatting
- ✅ Email notifications (rate-limited)
- ✅ Record ID generation (TR-XXXXXXXX format)

### What Needs to Be Added
- ❌ Three accountability text questions (Coverage, Gaps, Plan)
- ❌ End-of-Shift Recap confirmation (Yes/No + follow-up)
- ❌ Checklist photo upload (file → Google Drive → URL stored in sheet)
- ❌ Character minimum validation on text questions
- ❌ Flag logic: Day 4/5 low score email, consecutive NI alert, recap NO flagging
- ❌ Photo Drive folder organization (Location/Trainee/Day)
- ❌ GM audit view improvements

### What Needs to Be Fixed (Code Cleanup)
- ❌ Hardcoded trainer rosters → move to config sheet or centralized object
- ❌ DriveApp HTML parsing for trainer data → replace with direct config
- ❌ 4 separate analytics functions each scanning all data → single pass
- ❌ `sendNotificationEmail()` called but undefined → fix reference
- ❌ Double-comma bug in fallback trainer array (line 224)
- ❌ Remove dead code and unused functions

---

## PRE-LAUNCH CHECKLIST

### Form Functionality
- [ ] All 5 scoring categories display with 1–5 radio buttons
- [ ] Score summary auto-calculates (total, percentage, performance level)
- [ ] 3 accountability questions display with character minimums enforced
- [ ] Recap confirmation (Yes/No) works; NO reveals follow-up field
- [ ] Photo upload field is required — form blocks submission without it
- [ ] Photo upload accepts JPG, PNG, HEIC; enforces 10MB limit
- [ ] Location dropdown filters trainer list correctly
- [ ] Training Day dropdown populates Day 1–5
- [ ] Shift auto-selects based on time of day
- [ ] Submission timestamp captured accurately
- [ ] Confirmation screen displays with Record ID after submission

### Data & Backend
- [ ] All submissions write to Training Records sheet correctly
- [ ] Photo files stored in Google Drive, accessible for audit
- [ ] Accountability question responses stored in correct columns
- [ ] Recap status stored (Yes/No + reason if No)
- [ ] Submissions filterable by location, trainer, trainee, day
- [ ] GM/Devin audit view functional
- [ ] Flag logic active:
  - [ ] Day 4/5 score < 75% triggers email to Devin
  - [ ] Recap = NO flagged in audit view
  - [ ] Consecutive Needs Improvement triggers flag

### Permissions
- [ ] Trainers can only submit (not edit or delete past submissions)
- [ ] GMs can view all submissions for their location
- [ ] Devin has full access across all locations

### End-to-End Test
- [ ] Submit test form with all fields — confirm data lands in Training Records
- [ ] Submit test form with Recap = NO — confirm follow-up field required
- [ ] Submit test form with Day 5 + low score — confirm email alert
- [ ] Attempt submit without photo — confirm form blocks
- [ ] Verify photo appears in Drive folder
- [ ] Verify PAID VALIDATION sheet tracks correctly

---

*The form is the digital record. The signed checklist is the physical record. Both are required. Neither replaces the other. No form + no photo = no incentive. No exceptions.*
