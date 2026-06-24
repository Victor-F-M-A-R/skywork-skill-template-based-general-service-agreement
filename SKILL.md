---
name: template-based-general-service-agreement
description: Use this skill when the user wants to create a professional general service agreement, master service agreement, service provider agreement, vendor service agreement, contractor service agreement, or SOW-based service contract from a template-backed Word-generation workflow. This skill guides service agreement drafting, document structure, Python-based Word template execution, quality checks and fixes, and source-grounded validation.
---

**Bundled Template Assets**

Use the **Skill directory** shown above as the base directory. All relative paths below resolve from that directory.

Bundled files:
- `./generate_general_service_agreement_template.py`
- `./general_service_agreement_template.docx`

When generating or revising the Word document, copy the bundled Python generator and `.docx` template from the Skill directory into the active workspace. Prefer editing and running the copied generator. Write a new Python generator only if the bundled generator cannot support the requested output.

**Target write language**: {{ product_language }}

For every writing artifact, all user-visible text MUST use {{ product_language }}

# General Service Agreement

Use this skill to draft a professional service agreement between a client/customer and a service provider. It is suitable for B2B services, operational services, agency work, implementation work, maintenance services, outsourced services, and MSA + SOW arrangements.

## 1. General Service Agreement Drafting Guide

Draft service agreements around four recurring failure points: party identity, service variables, service obligations, and risk allocation.

### Party Metadata Consistency

Keep client/provider identity consistent across the title, recitals, party clause, SOWs, notices, invoices, schedules, and signature page. State legal names, roles, notice/billing/project contacts, representatives, signatories, covered affiliates/subcontractors/personnel/client group companies, effective date, service start date, SOW dates, renewal dates, and termination dates. Use `[TBD]` for missing entity data.

### Core Service Variable Consistency

Keep service variables consistent between the main agreement and SOWs/schedules: scope, excluded services, assumptions, dependencies, service levels, deliverables, acceptance criteria, fees, currency, invoicing, payment deadline, expenses, taxes, deposits, milestones, disputed invoices, SOW term, renewal, change orders, subcontractors, IP ownership, provider background materials, client materials, third-party materials, data/security terms, insurance, and termination notice. State order of precedence where SOWs vary from the master terms.

### Services, Deliverables, and Acceptance Clarity

Avoid vague commitments such as "provide services as needed", "support the client", or "perform all necessary work" unless scope and limits are clear. For each material obligation, state responsible party, service/deliverable/format, milestone, dependency, performance standard, acceptance/rejection process, correction process, and payment consequence.

### Risk Allocation

Allocate service-specific risks: scope creep, client dependencies, delayed approvals, change control, missed milestones, provider personnel/subcontractors, independent contractor status, no agency/employment/authority to bind, client materials, provider background IP, third-party materials, new deliverables, data, work product, warranties, indemnities, liability cap, confidentiality, data security, insurance, compliance, termination, transition assistance, return of materials, final invoices, work-in-progress, SOW wind-down, and survival.

### Document Style

By default, use a plain legal-document style: black text on a white background. Do not use colored headings, shaded backgrounds, decorative colors, or non-black accent colors unless the user explicitly says otherwise.

## 2. Document Structure

**Priority:** If the user gives a required structure, clause order, or template, follow the user's requirement first unless it creates a direct contract inconsistency.

Keep the structure clear enough to separate general service terms from project-specific scope, deliverables, fees, and acceptance criteria.

### First Decision

Choose one model before drafting:

- One-off service agreement: one fixed engagement with defined services, deliverables, fees, and acceptance.
- Master service agreement: repeat projects under separate SOWs; define SOW approval and precedence.
- Managed/ongoing service agreement: recurring services with service levels, reporting, renewal, transition, and data/security terms.
- Implementation/project service agreement: milestone-based work with dependencies, acceptance, change control, and delay consequences.

### Core Modules

Use these modules in order unless the user's source document requires a different order:

```text
Title
Parties
Background
Definitions
Appointment and Scope of Services
Statements of Work if applicable
Client Responsibilities
Provider Responsibilities
Deliverables and Acceptance
Change Control
Fees, Expenses, Invoicing, Taxes, and Disputed Invoices
Term, Renewal, and SOW Term if applicable
Confidentiality
Intellectual Property and Work Product
Data, Security, and Compliance if applicable
Warranties, Indemnity, Insurance, and Limitation of Liability
Termination and Effect of Termination
Dispute Resolution
Notices and Miscellaneous
Schedules
Signatures
```

### Do Not Miss

- Define included services, excluded services, deliverables, assumptions, dependencies, service levels, milestones, and acceptance criteria.
- Tie payment to fees, milestones, accepted deliverables, invoicing rules, taxes, expenses, and disputed invoice process.
- Put operational detail in schedules: SOW, fee schedule, acceptance criteria/SLA, data/security terms, and insurance requirements.
- Never leave background IP, deliverable ownership, change control, deemed acceptance, liability cap, or termination consequences undefined.

## 3. Template Execution

Use this template as the drafting blueprint. First identify the service agreement subtype, then map the user's facts and source materials into the Core Modules without changing the commercial structure unless the user requests a rewrite.

Use the Python generator listed in **Bundled Template Assets** as the source asset. Copy the generator and `.docx` template into the active workspace, prefer editing the copied generator, run it from the workspace, and verify the regenerated Word file. Write a new Python generator only if the bundled generator cannot support the requested output.

Execution rules:

- Use `[TBD]` for missing party names, dates, fees, scope, deliverables, acceptance criteria, insurance limits, signatories, governing law, and venue.
- Preserve user-provided SOWs, quotes, proposals, client comments, marked drafts, and prior service agreements unless instructed otherwise.
- Put variable-heavy terms in clauses, schedules, or SOW fields instead of burying them in prose.
- Keep title, service scope, SOW structure, acceptance language, fee schedule, signature blocks, and formatting consistent with the selected service agreement subtype.

## 4. General Service Agreement Quality Check and Fixes

Run both validation layers every time: mandatory service-agreement quality review and mandatory self-contained/source-grounded validation.

### Mandatory General Service Agreement Quality Check

Create `quality_check.md` beside the generated document. The file must summarize issues, severity, fixes, whether fixes were applied, and any remaining `[TBD]` items.

Check four areas:

- **Party metadata:** names, roles, addresses, tax/registration IDs, representatives, notice/billing contacts, signatories, contract numbers, SOW numbers, dates, SOWs, schedules, invoices if referenced, signature page, and file name if available.
- **Core service variables:** scope, excluded services, assumptions, dependencies, deliverables, service levels, acceptance criteria, term, start/milestone dates, renewal, termination notice, fees, invoicing, payment deadline, expenses, taxes, disputed invoices, change orders, subcontractors, insurance, IP, data/security, confidentiality, and transition obligations.
- **Services and acceptance:** each material obligation must identify responsible party, service/deliverable/standard, timing, acceptance/approval, correction process, and failure/payment consequence.
- **Risk allocation:** scope creep, SOW precedence, change control, unlimited service obligations, client delays, deemed acceptance, client materials, provider background IP, deliverables, data, confidentiality, data security, insurance, warranty, indemnity, liability cap, termination, transition assistance, final payment, work-in-progress, return of materials, and survival.

If fixable issues exist, edit the copied Word-generation `.py` file, regenerate the `.docx`, and update `quality_check.md`.

Use this `quality_check.md` structure:

```markdown
# General Service Agreement Quality Check

## Checks
- Party metadata: OK / Fixed / Issues / TBD
- Service variables: OK / Fixed / Issues / TBD
- Services, deliverables, and acceptance: OK / Fixed / Issues / TBD
- Risk allocation: OK / Fixed / Issues / TBD

## Material Risks
- Clause/Area; risk; current allocation; severity; fix applied.

## Remaining Fixes Needed
- [TBD]
```

### Mandatory Self-Contained And Source-Grounded Validation

Always call `validate_document` through `jupyter_execute` after generating or editing the final document. Include every available user-uploaded or source file; if no source file exists, run the validation with `source_files=[]` to check self-contained completeness against the user instruction.

```python
import sys
sys.path.insert(0, "/app/skills/common/doc_new_split_agent")
from validate import validate_document

result = await validate_document(
    generated_path="uploaded_files/output.docx",
    source_files=[
        {
            "path": "parsed_files/source.md",
            "origin": "uploaded",
            "description": "Parsed text of the user-uploaded source document",
        }
    ],
    user_instruction="""{verbatim user request}""",
)

print(result["verdict"])
print(result["self_contained"]["report"])
if result["source_grounded"]:
    print(result["source_grounded"]["report"])
```
