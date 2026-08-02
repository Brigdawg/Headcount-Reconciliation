# Headcount Control Tower

BambooHR interview case study — interactive prototype + methodology deck for **headcount reconciliation**.

## Live demo

After you push this repo to GitHub, the fastest public link is **Vercel** (free):

```bash
# from this folder
npm install
npx vercel --yes
```

Or connect the GitHub repo at [vercel.com/new](https://vercel.com/new) → import project → Framework: Vite → Deploy.  
You will get a URL like `https://headcount-control-tower.vercel.app`.

### GitHub Pages (alternative)

```bash
npm run build
# then enable Pages on the `dist/` folder, or use the `gh-pages` branch workflow
```

## Deliverables

| Artifact | Location |
|----------|----------|
| Interactive prototype | this app (`npm run dev` or live Vercel URL) |
| PowerPoint deck | [`deck/Headcount_Control_Tower_Case_Study.pptx`](deck/Headcount_Control_Tower_Case_Study.pptx) |

### Deck story arc

1. Problem  
2. Why it matters (stakes + who benefits)  
3. Evidence (finance quotes + market data)  
4. Hypotheses  
5. Interviews  
6. Product process  
7. Iterations (screenshots)  
8. Success metrics + illustrative impact  
9. Conclusion  

## Local run

```bash
npm install
npm run dev
```

## Regenerate the deck

```bash
pip3 install python-pptx pillow
python3 deck/generate_deck.py
```

## Notes

- Demo data is fictional (**Northline Systems**), patterned on a real multi-tab recon workflow  
- Market stats on slides are sourced (AFP, Kinnect/Gartner cite, IJIRMPS, etc.) — see speaker notes  
- Illustrative impact slide is labeled as a planning model, not a customer ROI claim  
