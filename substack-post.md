# The New Analyst: What AI Actually Changes About Financial Work

*A reflection on building an AI-powered underwriting tool — and what it revealed about the future of finance*

---

## The Problem Worth Naming

Real estate investment banking runs on Excel.

That is not a criticism. Excel is powerful, flexible, and universally understood. But it has a ceiling. A single analyst building a deal model from scratch — pulling market data, running sensitivity analysis, stress-testing assumptions — can expect to spend hours on work that is largely mechanical. The thinking that matters, the judgment about whether a deal is actually good, gets compressed into whatever time is left over.

This is the gap AI can close. Not by replacing the analyst. By freeing them to do the part of the job that only humans can do.

---

## Structure: What the Tool Does

Over the past several months, as part of MGMT 69000 at Purdue, I built an AI-powered commercial real estate underwriting platform using Python, Streamlit, and a set of integrated data APIs.

The tool does the following:

- **Financial modeling** — IRR, DSCR, equity multiple, amortization schedule, and before-tax cash flow across a full hold period
- **Value-add renovation modeling** — spreads capex across construction years, applies post-completion rent bumps, calculates payback period and pre/post NOI impact
- **Correlated Monte Carlo simulation** — 10,000 scenarios with mathematically correlated shocks across rent growth, occupancy, cap rate, and expense variables
- **Machine learning valuation** — gradient boosting model trained on 300 NCREIF-calibrated transaction records, outputs an UNDERVALUED / FAIR / OVERVALUED assessment
- **Live market data** — FRED, Census ACS, BLS, HUD Fair Market Rents, RentCast, and others, pulled in parallel and synthesized into a single market context score
- **LP/GP waterfall** — two-tier promote structure, distributable cash calculated at each hurdle
- **Recommendation engine** — buy/hold/pass signal built from seven weighted factors
- **CI/CD pipeline** — 157 automated tests, lint enforcement, continuous deployment to Streamlit Community Cloud

The tool is live. The code is public. It runs in a browser.

---

## Purpose: Why Build This

The purpose was not to replace a financial model. It was to compress the distance between a deal and a decision.

In traditional underwriting, a junior analyst might spend a full day building a model before anyone can have a meaningful conversation about whether the deal makes sense. The assumptions are debated. The data is gathered manually. The sensitivity analysis is run in a separate tab. The market context lives in someone's head or in a separate research brief.

This tool collapses that pipeline. You enter a deal — address, units, purchase price, financing terms — and in under a minute you have financial projections, market context, a risk distribution, an ML-based valuation check, and a recommendation. Not a final answer. A starting point that is much further along than a blank spreadsheet.

The purpose is speed-to-judgment. Not speed for its own sake — but getting the analytical scaffolding out of the way so the actual thinking can begin sooner.

---

## Utility: What a Finance Professional Can Take From This

Whether or not you build something like this, there are transferable principles here.

**APIs change what market research costs.** FRED, Census, BLS, HUD — these are free, well-documented, and available right now. The barrier to pulling live market data into a model is no longer cost or access. It is knowing that the APIs exist and writing thirty lines of code to use them.

**Monte Carlo simulation is not exotic anymore.** Running 10,000 correlated scenarios used to require specialized software. With NumPy and SciPy, it is about fifty lines of Python. The math has not changed. The barrier to implementing it has.

**The discipline of validation is what makes AI-assisted work credible.** Every feature in this tool has tests. Every output that comes from synthetic data is labeled as such. The CI pipeline fails if the lint check fails, before any code can reach production. This is what makes it possible to trust what the tool outputs. Anyone using AI to build analytical tools for professional use needs this discipline, or the tool is not credible.

**AI is a force multiplier, not a decision-maker.** The tool does not tell you whether to buy a building. It tells you what the numbers look like under a range of assumptions, what the market says, and what comparable transactions suggest about value. The investment judgment stays with the human. AI accelerates the analysis. It does not substitute for the decision.

---

## Meaning: What This Represents

The question I kept coming back to while building this was not whether AI can do this — it clearly can. The question was what it means that AI can do this.

The role of a junior analyst in finance has always mixed two kinds of work: mechanical work (building models, pulling data, formatting outputs) and judgment work (synthesizing information, advising clients, making recommendations). AI is now capable of doing the mechanical work faster and more consistently than a person. That is not a threat to the profession. It is a reallocation of time toward the work that is actually hard.

The analysts who will do well in the next decade are not the ones who resist these tools. They are the ones who understand them well enough to direct them — to define the right problem, evaluate the output critically, catch the errors the model cannot catch, and make the decision that requires human judgment.

This project gave me a concrete experience of what that looks like. The AI wrote the code. I decided what to build, evaluated every output, caught the gaps, and made the product decisions. The AI was faster at implementation than I will ever be. I was better at knowing what was worth building.

That division of labor — AI as instrument, human as driver — is not a compromise. It is what makes both more effective.

---

## Where the Tool Is

The tool is live at: *[Streamlit URL — add before publishing]*

The code is on GitHub: [github.com/peterjohn1298/re-underwriting-tool](https://github.com/peterjohn1298/re-underwriting-tool)

If you are in real estate, finance, or building AI-assisted tools of your own, the repository is open. The methodology is documented. The tests run on every push.

---

*Peter John | MGMT 69000: Mastering AI for Finance | Purdue University*
