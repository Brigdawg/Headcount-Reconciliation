# Headcount Control Tower

An interactive product prototype for **headcount reconciliation** — bringing Finance and HR onto one shared view of plan vs. actual headcount, variance drivers, and hiring approvals.

Built as a BambooHR interview case study. Demo data is fictional (**Northline Systems**).

**Live demo:** [https://headcount-reconciliation.vercel.app](https://headcount-reconciliation.vercel.app)

## What’s inside

**Interactive prototype** — a BambooHR-styled app with:

- Story KPIs and headcount outlook (H2, YTD, 12-month views)
- Variance bridge explaining plan vs. actual gaps
- Department budgets and employment-type filters
- Approvals queue, roles table, and audit/snapshot history
- Finance vs. HR persona views
- Excel export preview (illustrative, not a real download)

**Case study deck** — process and research narrative:  
[`deck/Headcount_Control_Tower_Case_Study.pptx`](deck/Headcount_Control_Tower_Case_Study.pptx)

## Try it locally

```bash
npm install
npm run dev
```

Then open the URL Vite prints (typically `http://127.0.0.1:5173`).

## Tech

Vite · React · TypeScript · Recharts
