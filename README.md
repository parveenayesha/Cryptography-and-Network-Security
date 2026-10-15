# CNS Smart Course Website — U23CS704

**Project title:** AI-Enabled Framework for Outcome-Based Smart Course Design, Automated CO–PO Mapping and Adaptive Content Generation for Cryptography and Network Security

This is a responsive static teaching website for the CNS syllabus supplied for the capstone. It uses only HTML, CSS and JavaScript; no build tools or external libraries are required.

## Files
- `index.html` — course overview, draft COs, all five syllabus units, 15-week schedule, assessment plan, AI policy, contact placeholders and interactive activities.
- `styles.css` — responsive styling, navigation, tables and cards.
- `script.js` — mobile navigation, Caesar cipher encrypt/decrypt, quiz scoring.

## Run locally
Open `index.html` in a current browser. For a local web server, run `python -m http.server 8000` in this folder and open `http://localhost:8000`.

## Publish with GitHub Pages (step by step)
1. Sign in at https://github.com/ or create an account.
2. Select **New repository** and create a public repository named `cns-smart-course` (or another name you prefer).
3. Upload `index.html`, `styles.css`, `script.js` and `README.md` to the repository root, then commit.
4. In the repository, open **Settings → Pages**.
5. Under **Build and deployment**, choose **Deploy from a branch**; choose branch `main` and folder `/(root)`; click **Save**.
6. Wait for the Pages workflow to finish. The expected URL pattern is `https://YOUR-USERNAME.github.io/cns-smart-course/` (replace `YOUR-USERNAME` with the account name).
7. Open the deployed URL and test all navigation, the cipher lab and quiz. Add the actual URL and a screenshot of the browser address bar to the portfolio.
8. Edit the contact placeholders before publication and commit the changes.

## Credentials and tokens
This is a static educational site. It has no login endpoint, authentication middleware, backend API, cookies, access tokens, refresh tokens or credentials. The quiz and cipher run only in the browser and do not send answers or entered text to a server. Do not add student data, passwords, API keys or GitHub personal access tokens to this repository. To publish, authenticate to GitHub in your browser and ensure the account has permission to commit to the repository. If automation is added later, use GitHub Actions secrets / environment secrets and least-privilege permissions; never hard-code tokens in HTML, CSS, JavaScript or commit history.

## Academic note
The supplied syllabus does not specify unit-wise lecture hours, official semester dates, detailed CIE subcomponent weights or official CO–PO mappings. These website elements are marked as provisional/draft and should be approved by the course coordinator before formal use. The official question-paper quality audit is pending the source CNS question paper; the portfolio includes a sample paper as a design illustration, not as a substitute for the source-paper audit.
