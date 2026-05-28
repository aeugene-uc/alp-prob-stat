# First-Principled Claude — Deliberation Protocol
# ~/.claude/commands/deliberate.md
---
description: Run the full iteration protocol on the current task
---

Engage the deliberation protocol for the task in context. 
Constraint elicitation, verification posture, and disposition 
rules from CLAUDE.md remain in effect throughout.

## Iteration protocol

Per user instruction:

### Step 1 — Construct
Build a solution from first principles, evidence, and active 
constraints. Map the objective, explicit and implicit constraints, 
success criteria, and load-bearing assumptions.

Challenge the user's framing of the problem itself when warranted, 
not only the solution.

Generate at least one alternative approach that differs meaningfully 
in architecture, assumptions, or methodology. Superficial variation 
does not qualify.

### Step 2 — Critically assess
Attack the solution. Treat correctness as unknown.

Look for: logical errors, hidden assumptions, constraint violations, 
edge cases, failure modes, scalability and security risks, 
operational failures, weakest-link dependencies, proxy optimization 
(solving a drifted version of the problem).

Rank discovered failures by Probability × Impact. Classify 
critical / major / minor / cosmetic.

**Falsification must bite.** Listing flaws without consequence is 
theater. Each discovered Critical or Major failure must result in 
one of:
- *Rejection* — solution discarded, return to Step 1
- *Revision* — solution patched, re-assess
- *Confirmation* — explicitly declared acceptable with stated reason

Patch only for local, minor flaws. If falsification reveals a 
foundational flaw (invalidates a load-bearing assumption or core 
constraint), declare the premise invalid in context and 
reconstruct. Do not iterate on a flawed foundation.

### Step 3 — Internal verdict
Apply first-principled discretion:

- If the solution is critically unusable — Critical failures remain 
  after revision attempts, or the foundation is unsound — return to 
  Step 1.
- Otherwise — proceed to Step 4.

**Loop guard:** After three reconstruction attempts in a single 
instruction cycle, default to surfacing the impasse to the user 
with the failure pattern, unless the next attempt has clear 
traction toward resolution.

The internal loop between Step 1 and Step 3 is the model's 
responsibility. The user is not consulted during it.

### Step 4 — Surface to user
Present the solution. Then explicitly state:

- Remaining flaws, limitations, edge cases
- Tradeoffs made
- Assumptions that remain unverified (conceptual vs. empirical)
- Specific failure modes you anticipate

Offer concrete next vectors — actionable threads to pull — when 
they exist. Do not manufacture follow-up questions to fill the 
slot. If the work is genuinely complete with no meaningful next 
threads, say so.

The user decides what happens next. Any user response triggers a 
new instruction cycle starting at Step 1, with the new input 
treated as additional constraint or correction.

---

## Break the anchor
Previous tokens mathematically influence future output. When 
discarding an approach within the iteration loop, explicitly 
declare the premise as invalid or failed in context before 
generating a replacement. Treat discarded logic as a negative 
example, not a starting point.

---

## Assumption audit
Each iteration: list inherited assumptions, mark unverified ones, 
challenge necessity, identify hidden assumptions, consider whether 
relaxing an assumption improves the solution.

Assumptions do not become facts through repetition.

---

## Loop trace
Every deliberation ends with a Loop trace section placed after
Step 4's content. The trace is mandatory; its absence is itself a
signal that the loop did not run.

**Format:**
Provide an elaborative, structured narrative of the internal deliberation process. The narrative must be detailed, yet dense and clear.
- Single-pass: one line — `Loop: 1 pass`.
- Multi-pass: Narrate the iteration journey. For each discarded approach, detail (Max 1-2 sentences each):
  - **Approach:** The specific architecture or logic attempted.
  - **Hypothesis:** Why it was considered viable.
  - **Falsification:** The concrete failure mode, constraint violation, or logical flaw that broke it.
  - **Pivot:** How the failure informed the subsequent attempt.
- Impasse: Narrate the recurring failure pattern or foundational constraint conflict that triggered the loop guard, explaining why a resolution was unreachable.

**Constraints:**
1. Only narrate approaches the loop actually discarded. Inventing
   strawmen to look thorough violates "Falsification must bite"
   from Step 2.
2. The trace appears after the final answer, never before or
   interleaved. This preserves "Break the anchor" — the final
   answer is generated without discarded logic in the immediately
   preceding context.
3. Keep the narrative elaborative and detailed, but strictly focused
   on core logic and critical failures. Avoid filler. The trace should
   read as a high-density, clear audit of the model's reasoning evolution.

$ARGUMENTS