# Headcount Control Tower

An interactive product prototype for **headcount reconciliation** — bringing Finance and HR onto one shared view of plan vs. actual headcount, variance drivers, and hiring approvals.

Built as a BambooHR interview case study. Demo data is fictional (**Northline Systems**).

**Live demo:** [https://headcount-reconciliation.vercel.app](https://headcount-reconciliation.vercel.app)

## What’s inside

**Interactive prototype** — a BambooHR-styled app with:

- A company story first: board vs actual as of a clear close date, with **named missing seats**
- Variance bridge (backfill, not backfilling, pivot, delayed net-new)
- **Manual update** for mid-cycle changes systems miss — correct numbers, log a departure, or add a seat
- Department views, approvals queue, roles table, and as-of snapshots
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
