# ACCOUNTABILITY FORM — SPECIFICATION & LAUNCH CHECKLIST
**Three Points Hospitality — C.O.R.E. Training System**
**Last Updated: March 16, 2026**

---

## PURPOSE

The Accountability Form is submitted by the **trainee** at the end of every training shift. It is the digital paper trail that runs alongside the signed physical checklist. Together, they confirm the shift happened, the trainer covered the material, and the trainee engaged with it.

This document defines the form spec, the three required questions, the photo upload requirement, and the pre-launch verification checklist.

---

## FORM OWNERSHIP

| Field | Detail |
|---|---|
| **Built and managed by** | Devin (app side) |
| **Submitted by** | Trainee — end of every training shift |
| **Reviewed by** | GM — weekly audit |
| **Storage** | App backend; exportable for GM review |

---

## THE THREE REQUIRED QUESTIONS

These are the three questions every trainee answers on every submission. They are non-negotiable and not role-specific — the same three questions appear for every role, every day.

### Question 1: Shift Confirmation
> **"What was the most important thing you practiced today?"**

- Field type: Open text
- Minimum 20 characters required (prevents single-word or blank submissions)
- Purpose: forces active recall of the shift's primary skill or topic
- GM flag trigger: if the answer has no connection to that day's checklist content, flag for trainer follow-up

### Question 2: Standard Self-Assessment
> **"On a scale of 1-5, how confident are you in executing what you practiced today without help?"**

- Field type: Radio button — 1 (Not at all confident) through 5 (Fully confident)
- Purpose: surfaces self-reported gaps before they become solo-shift problems
- GM flag trigger: score of 1 or 2 on any Day 4 or Day 5 submission triggers a manager check-in before the next shift

### Question 3: Trainer Confirmation
> **"Did your trainer complete the End-of-Shift Recap with you before you clocked out?"**

- Field type: Yes / No toggle
- If NO: follow-up field appears — "What did you do instead of the recap?"
- Purpose: enforces the non-negotiable recap block from the floor checklist
- GM flag trigger: any NO response triggers a trainer accountability review

---

## PHOTO UPLOAD REQUIREMENT

- Field label: "Upload a photo of your signed Day [X] checklist"
- Required field — form cannot be submitted without a photo
- Accepted formats: JPG, PNG, HEIC (mobile-first)
- File size limit: 10MB
- Photo is reviewed in the GM's weekly audit
- If photo is illegible or unsigned: GM follows up with trainer same day

---

## METADATA CAPTURED (AUTO)

The following fields are auto-populated — the trainee does not fill these in:

| Field | Source |
|---|---|
| Trainee name | Account login |
| Training day number | Selected by trainee at form open |
| Role | Account profile |
| Location | Account profile |
| Submission timestamp | System |
| Trainer name | Selected by trainee at form open |

---

## FORM FLOW

| Step | Action |
|---|---|
| 1 | Form opens |
| 2 | Select: Training Day (Day 1 through Day 5) |
| 3 | Select: Trainer Name (dropdown from active trainer list) |
| 4 | Question 1 — Open text (20 char minimum) |
| 5 | Question 2 — Confidence rating (1-5) |
| 6 | Question 3 — Recap confirmation (Yes / No); if No, follow-up text field required |
| 7 | Photo Upload — Required |
| 8 | Submit |
| 9 | Confirmation screen: "Submitted. See you tomorrow." |

---

## PRE-LAUNCH VERIFICATION CHECKLIST

Complete every item before the first training cohort begins.

### Form Functionality
- [ ] All three questions are live and display correctly on mobile
- [ ] Question 1 text field enforces 20-character minimum
- [ ] Question 2 radio buttons are mutually exclusive (cannot select two)
- [ ] Question 3 YES/NO toggle works; NO response reveals follow-up field
- [ ] Follow-up field is required if NO is selected on Question 3
- [ ] Photo upload field is required — form cannot submit without it
- [ ] Photo upload accepts JPG, PNG, and HEIC formats
- [ ] File size limit (10MB) is enforced with a clear error message
- [ ] Training Day dropdown populates correctly (Day 1 through Day 5)
- [ ] Trainer Name dropdown pulls from active trainer list
- [ ] Submission timestamp is captured accurately
- [ ] Confirmation screen displays after successful submission

### Data & Storage
- [ ] All submissions route to the correct backend / data store
- [ ] Photo files are stored and accessible for GM review
- [ ] Trainee name, role, and location populate from account profile correctly
- [ ] Submissions are filterable by location, role, and training day
- [ ] GM audit view is functional — can see all submissions with metadata
- [ ] Flag logic is active:
  - [ ] Question 2 score of 1 or 2 on Day 4 or Day 5 triggers manager alert
  - [ ] Question 3 NO response triggers trainer accountability review

### Access & Permissions
- [ ] Trainees can only submit for their own account
- [ ] Trainees cannot edit or delete past submissions
- [ ] GMs can view all submissions across their location
- [ ] Trainers can view their own trainees' submissions only
- [ ] No one below GM can view submissions from other trainers' trainees

### End-to-End Test
- [ ] Submit a test form as a trainee with all fields completed
- [ ] Verify photo uploads and is visible in the GM view
- [ ] Verify metadata (name, role, location, timestamp) populated correctly
- [ ] Submit a test form with Question 3 = NO — confirm follow-up field appeared
- [ ] Submit a test form with Question 2 = 1 on Day 5 — confirm GM alert triggered
- [ ] Attempt to submit without a photo — confirm form blocks submission
- [ ] Confirm confirmation screen displays after each successful test

### Communication
- [ ] All trainees know they are required to submit the form before clocking out
- [ ] All trainers know form submission is required before their incentive is counted
- [ ] GMs know the audit cadence (weekly) and flag protocols
- [ ] Support contact is defined if form has a technical issue during a shift

---

## GM WEEKLY AUDIT PROTOCOL

Every Monday, the GM reviews the prior week's submissions:

1. Pull all submissions from the past 7 days
2. Check for missing submissions (training days with no form filed)
3. Review all flagged entries (low confidence scores, NO recap responses)
4. Check photo quality — confirm checklists are signed and legible
5. Follow up with trainer on any missing or flagged entries within 24 hours
6. Log audit completion in the weekly manager report

> **Target: 100% submission rate.** Any week below 90% triggers a process review with the training lead.

---

## INCOMPLETE SUBMISSION HANDLING

| Situation | Action |
|---|---|
| Trainee did not submit | GM follows up with trainer same day; shift flagged as incomplete |
| Photo missing or illegible | Trainer resubmits corrected photo within 24 hours |
| Question 3 = NO | GM reviews with trainer; escalates if pattern repeats |
| Q2 score 1-2 on Day 4 or Day 5 | Manager check-in with trainee before next shift |
| Trainer name not selected | Form is invalid; trainee resubmits |

---

*The form is the digital record. The signed checklist is the physical record. Both are required. Neither replaces the other.*