# AI-Enabled Framework for Outcome-Based Smart Course Design, Automated CO–PO Mapping and Adaptive Content Generation for Cryptography and Network Security

Course: U23CS704 — Cryptography and Network Security

## 1. Project files

- `CNS_AI_Enabled_Smart_Course_Portfolio.docx` and `.pdf`: main FDP capstone portfolio.
- `CNS_Sample_Assessment_Question_Paper.docx` and `.pdf`: instructor-created sample practice assessment, labelled as a draft.
- `CNS_OBE_Attainment_and_Assessment_Workbook.xlsx`: illustrative OBE dashboard, CO–PO draft matrix, attendance, marks, assessment blueprint, exit-ticket/reflection tabs and formula-based attainment demo.
- `website/`: complete static website source (`index.html`, `styles.css`, `script.js`).
- `automation/CNS_Course_Automation.gs`: Google Apps Script starter.
- `docs/google_forms_setup.md`: setup notes for instructor-owned Google Forms/Sheets.
- `docs/authentication_notes.md`: authentication, request-flow and credential/token handling notes.
- `research/research_integration.md`: five recent research summaries and a future research direction.
- `assessment/CNS_Sample_Assessment.md`: editable Markdown copy of the sample assessment.

## 2. Run the site locally

Open `website/index.html` in a modern browser. Alternatively, open a terminal in the `website` folder and run:

```bash
python -m http.server 8000
```

Then open `http://localhost:8000`.

## 3. Publish with GitHub Pages

1. Sign in at [GitHub](https://github.com/).
2. Create a **public** repository named `cns-smart-course` (or another approved name).
3. Upload `index.html`, `styles.css`, `script.js`, and `README.md` from the `website/` folder to the repository root and commit them to the `main` branch.
4. Open **Settings → Pages**.
5. Under **Build and deployment**, choose **Deploy from a branch**, select `main` and `/(root)`, then save.
6. Wait for the Pages deployment to finish.
7. The expected address pattern is `https://YOUR-USERNAME.github.io/cns-smart-course/`. Replace `YOUR-USERNAME` with the publishing account's actual GitHub username.
8. Test the live website and replace contact placeholders. Add the real URL and screenshots to the final submission only after publication is confirmed.

Documentation: https://docs.github.com/en/pages

## 4. Authentication and credential flow

The generated CNS website is static. A learner's browser requests the HTML, CSS and JavaScript from GitHub Pages; cipher and quiz interactions are executed locally in the browser. The website has no login endpoint, authentication middleware, backend API, database, session, cookie, access token or refresh token. It does not send cipher text or quiz answers to a server.

Publishing is a separate account-level action: the person publishing signs into GitHub and must have repository write permission to commit. A public GitHub Pages website ordinarily does not require learner sign-in. Never include a password, personal access token, API key or student record in any public website file or screenshot. If a future GitHub Actions workflow needs a secret, store it in approved repository/environment secrets and use least-privilege permissions.

No existing external repository URL was supplied, and no live repository or Pages deployment is claimed in this package. The website source is provided so it can be published from the author's account.

## 5. Academic assumptions and pending items

- The official syllabus specifies the course title/code, prerequisites, three credits, 40 CIE marks, 60 SEE marks, official COs and five units. It does not specify unit-wise teaching hours, a formal CO–PO/PSO matrix, semester dates, detailed CIE component weights or attainment thresholds. Those entries are identified as proposed examples.
- Workbook marks and attainment percentages are synthetic demonstrations, not actual student performance.
- The 50-mark sample paper is a newly authored practice assessment; its duration and pattern are assumed and must not be presented as the official SEE format.
- The `blooms-quality-and-rebalance-report` skill requires the original CNS question paper before its question-wise quality audit and marked rebalanced paper can be produced. That specific formal audit remains pending the actual paper.
- Replace all faculty/institution/contact placeholders and verify all mapped outcomes, assessment rules, research citations and formulas before submission.
