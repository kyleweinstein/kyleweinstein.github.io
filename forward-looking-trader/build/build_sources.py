"""Builds Source_Register.xlsx"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
F="Arial"; HDR=PatternFill("solid",fgColor="1F3A5F"); thin=Side(style="thin",color="BBBBBB"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
def sheet(wb,name,title,note,heads,widths,rows,first=False):
    ws=wb.active if first else wb.create_sheet(name); ws.title=name
    ws["A1"]=title; ws["A1"].font=Font(name=F,size=14,bold=True,color="1F3A5F"); ws["A2"]=note; ws["A2"].font=Font(name=F,size=9,italic=True,color="555555")
    for i,(h,w) in enumerate(zip(heads,widths),1):
        c=ws.cell(row=4,column=i,value=h); c.font=Font(name=F,size=10,bold=True,color="FFFFFF"); c.fill=HDR; c.border=BOX; c.alignment=Alignment(wrap_text=True,vertical="center")
        ws.column_dimensions[c.column_letter].width=w
    for r,row in enumerate(rows,5):
        for i,v in enumerate(row,1):
            c=ws.cell(row=r,column=i,value=v); c.font=Font(name=F,size=10); c.border=BOX; c.alignment=Alignment(wrap_text=True,vertical="top")
    ws.freeze_panes="B5"; ws.sheet_view.showGridLines=False; return ws
E="https://www.sec.gov/Archives/edgar/data/723125/"
def mu(i,form,period,filed,acc,doc,claims):
    return (i,f"Micron Technology, Inc., Form {form}, {period}","Primary — SEC filing",filed,E+acc.replace("-","")+"/"+doc,claims,"L1 §B, D, E; Registry R-rows",
            "Verified (primary data)","Figures retrieved 16 Nov 2026 through an SEC-filings data service (MetricDuck), each tied to this accession number.",f"Accession {acc}. Values are latest-filed and may reflect later recasts of the period.")
V1="Corroborated (secondary only)"; NV="Not verified this session"; BIB="Bibliographic — from memory; publisher page not reachable"
src=[
 ("S1","Norges Bank Investment Management, In Good Company (podcast, host Nicolai Tangen), “Stan Druckenmiller: Inside the mind of a legendary investor”","Primary — interview","6 Nov 2024 (highlights version 8 Nov 2024)",
  "https://www.nbim.no/en/news-and-insights/podcast/2024/stan-druckenmiller-inside-the-mind-of-a-legendary-investor/","Looking ~18–24 months ahead rather than at the present; cyclical capacity-closure example.","Blueprint §1.1, §9; L1 §C1, K",V1,
  "Web search 16 Nov 2026: episode title/date confirmed by several podcast indexes; the 18–24-month and capacity passages confirmed in secondary transcript excerpts (S7).","nbim.no is outside the build environment's allowed network hosts, so the primary page and audio were NOT opened. No verbatim quotation used."),
 ("S2","Morgan Stanley, Hard Lessons: “Stan Druckenmiller: Invest, Then Investigate” (with Iliana Bouzali)","Primary — interview / transcript","Published 27 Feb 2026; recorded 30 Jan 2026",
  "https://www.morganstanley.com/insights/videos/hard-lessons/duquesne-stan-druckenmiller-iliana-bouzali","Trades conceptualised as an ~18-month-to-3-year evolution (some ~1 year, some up to 5); acting decisively and changing course quickly; Teva example; views on contrarianism and technical signals.","Blueprint §1.1, §9; L1 §C1–C2, K",V1,
  "Web search 16 Nov 2026: dates and participants confirmed; the 18-month-to-three-year passage appeared in search excerpts of the Morgan Stanley page and in a third-party transcript (S8).","morganstanley.com not reachable from the build environment; transcript not read in full. The technical-signals caution was NOT found (see claim C-06)."),
 ("S3","Edward Chancellor (ed.), Capital Returns: Investing Through the Capital Cycle — A Money Manager’s Reports 2002–15. Palgrave Macmillan, 2015.","Book","2015","https://www.edwardchancellor.com/books/capital-returns","Capital-cycle approach; industry supply response.","Blueprint §4 (Wk 6), §9",BIB,"Standard bibliographic details.","Author page not reachable; details to be re-checked against the publisher record."),
 ("S4","Michael J. Mauboussin & Alfred Rappaport, Expectations Investing: Reading Stock Prices for Better Returns. Revised and updated edition, Columbia Business School Publishing, 2021 (first edition Harvard Business School Press, 2001).","Book","2021 / 2001","https://www.expectationsinvesting.com/","Reverse-engineering price-implied expectations; revisions to expectations.","Blueprint §4 (Wk 11), §9",BIB,"Standard bibliographic details.","Book site not reachable; edition details to be re-checked."),
 ("S5","Stanley Druckenmiller, talk at the Lost Tree Club, North Palm Beach, Florida (with Q&A led by Ken Langone)","Secondary — circulating transcripts of a primary talk","18 Jan 2015","Transcriptions: gist.github.com/bluechoochoo/a034da52c64ac6fcb637 ; aryadeniz.substack.com/p/stanley-druckenmillers-lost-tree ; theideafarm.com/miscellaneous/speech-at-lost-tree-club/",
  "“Never, ever invest in the present”; visualise the situation ~18 months ahead; early-2000 technology trade (sold in January, bought ~$6bn, lost ~$3bn in about six weeks).","L1 §C1, F5, K; Blueprint Wk 3, 19",V1,
  "Web search 16 Nov 2026: the same passages appear consistently across several independent transcriptions.","No official transcript exists; transcriptions note conversion errors. Paraphrased in the lesson; the short phrase is attributed 'as transcribed'."),
 ("S6","Morgan Stanley, promotional post on X for the Hard Lessons Druckenmiller episode","Primary — issuer's own description","27 Feb 2026","https://x.com/MorganStanley/status/2027444279593890010","Episode covers 'why contrarianism is overrated'.","L1 §C1, F4",V1,"Appeared in web search results 16 Nov 2026.","Describes the episode; the exact spoken sentence was reported by third parties."),
 ("S7","Secondary coverage of S1: Marcellus, “Long read: Stan Druckenmiller: Inside the mind of a legendary investor”; A Letter a Day, Letter #235","Secondary — transcript / excerpt","Nov 2024","https://marcellus.in/story/long-read-stan-druckenmiller-inside-the-mind-of-a-legendary-investor/ ; https://aletteraday.substack.com/p/letter-235-stan-druckenmiller-and","Carries the 18–24-month and capacity-closure passages from S1.","Supports S1",V1,"Search-result summaries only.","Pages themselves not opened (network)."),
 ("S8","Secondary coverage of S2: A Letter a Day, Letter #320; third-party auto-generated transcripts","Secondary — transcript","Feb–Mar 2026","https://aletteraday.substack.com/p/letter-320-stan-druckenmiller-and","Carries the 18-month-to-3-year passage and the Teva figures from S2.","Supports S2",V1,"Search-result summaries only.","Auto-generated transcripts; wording unreliable."),
 mu("M1","10-K","fiscal year ended 1 Sep 2022","Oct 2022 (exact date not recorded)","0000723125-22-000048","mu-20220901.htm","FQ4 FY22 revenue $6.64B, gross margin 39.5%, net income $1.49B, capex $3.61B; FY22 quarterly history."),
 mu("M2","10-Q","quarter ended 1 Dec 2022 (FQ1 FY23)","22 Dec 2022","0000723125-22-000072","mu-20221201.htm","FQ1 FY23 revenue $4.09B, GM 21.9%, net loss $0.20B, capex $2.45B. MD&A: industry 'deteriorated sharply'; profitability 'challenged throughout calendar 2023'; 'reducing wafer starts and capital expenditures'; ~10% headcount reduction; revenue −39% q/q, −47% y/y."),
 mu("M3","10-Q","quarter ended 2 Mar 2023 (FQ2 FY23)","29 Mar 2023","0000723125-23-000022","mu-20230302.htm","Revenue $3.69B, GM −32.7%, net loss $2.31B, capex $2.21B."),
 mu("M4","10-Q","quarter ended 1 Jun 2023 (FQ3 FY23)","29 Jun 2023","0000723125-23-000041","mu-20230601.htm","Revenue $3.75B, GM −17.8%, net loss $1.90B, capex $1.56B."),
 mu("M5","10-K","fiscal year ended 31 Aug 2023","Oct 2023 (exact date not recorded)","0000723125-23-000054","mu-20230831.htm","FQ4 FY23 revenue $4.01B, GM −10.8%, net loss $1.43B, capex $1.46B. FY23 sums: revenue $15.5B, net loss $5.8B, capex $7.7B."),
 mu("M6","10-Q","quarter ended 30 Nov 2023 (FQ1 FY24)","Dec 2023/Jan 2024 (exact date not recorded)","0000723125-23-000077","mu-20231130.htm","Revenue $4.73B, GM −0.7%, net loss $1.23B, capex $1.80B."),
 mu("M7","10-Q","quarter ended 29 Feb 2024 (FQ2 FY24)","Mar 2024 (exact date not recorded)","0000723125-24-000007","mu-20240229.htm","Revenue $5.82B, GM 18.5%, net income $0.79B, capex $1.38B."),
 mu("M8","10-Q","quarter ended 30 May 2024 (FQ3 FY24)","Jun 2024 (exact date not recorded)","0000723125-24-000019","mu-20240530.htm","Revenue $6.81B, GM 26.9%, net income $0.33B, capex $2.09B."),
 mu("M9","10-K","fiscal year ended 29 Aug 2024","Oct 2024 (exact date not recorded)","0000723125-24-000027","mu-20240829.htm","FQ4 FY24 revenue $7.75B, GM 35.3%, net income $0.89B, capex $3.12B."),
 mu("M10","10-Q","quarter ended 28 Nov 2024 (FQ1 FY25)","Dec 2024/Jan 2025 (exact date not recorded)","0000723125-24-000047","mu-20241128.htm","Revenue $8.71B, GM 38.4%, net income $1.87B, capex $3.21B."),
 mu("M11","10-K","fiscal year ended 3 Sep 2026","Oct 2026 (exact date not recorded)","0000723125-26-000023","mu-20260903.htm","FQ4 FY26 revenue $54.23B, GM 86.8%, net income $37.70B, capex $11.11B."),
 ("M12","Micron Technology, Inc., Forms 10-Q for FQ1–FQ3 FY2026 (quarters ended 27 Nov 2025, 26 Feb 2026, 28 May 2026)","Primary — SEC filing","Dec 2025 – Jun 2026",E+"000072312525000046/mu-20251127.htm ; "+E+"000072312526000006/mu-20260226.htm ; "+E+"000072312526000015/mu-20260528.htm",
  "Quarterly revenue $13.64B / $23.86B / $41.46B; net income $5.24B / $13.79B / $28.24B; capex $5.39B / $6.39B / $7.83B. With M11: FY26 revenue $133.2B, net income $85.0B, capex $30.7B.","L1 §E","Verified (primary data)","Retrieved 16 Nov 2026 via SEC-filings data service.","Accessions 0000723125-25-000046, -26-000006, -26-000015. Fiscal-year sums are this curriculum's arithmetic."),
 ("M13","End-of-day share prices for MU (IEX historical data) and period-end shares outstanding, via MetricDuck","Market-data feed (not SEC-filed)","Retrieved 16 Nov 2026","IEX HIST via MetricDuck; terms: https://www.iex.io/legal/hist-data-terms",
  "Closes: $49.44 (22 Dec 2022), $63.56 (29 Mar 2023), $139.12 (24 Jun 2024), $89.24 (23 Dec 2024), $1,074.71 (2 Oct 2026), $798.51 (12 Nov 2026). Shares 1.091B (1 Dec 2022), 1.131B (3 Sep 2026).","L1 §D, E","Verified (single feed)","Retrieved 16 Nov 2026.","Price-only returns, not total returns. Market values and trailing P/E (8.7× and 10.6×) are this curriculum's arithmetic."),
 ("B1","Glenn W. Brier, “Verification of Forecasts Expressed in Terms of Probability,” Monthly Weather Review 78(1), 1950.","Journal article","1950","—","Brier score.","L1 Practice 4; Registry",BIB,"Standard reference.","Workbook uses the common one-sided binary form (range 0–1)."),
 ("B2","Allan H. Murphy, “A New Vector Partition of the Probability Score,” Journal of Applied Meteorology 12(4), 1973.","Journal article","1973","—","Reliability–resolution–uncertainty decomposition.","Registry → Calibration",BIB,"Standard reference.","Exact only with identical forecasts within bins; workbook reports the residual."),
]
wb=Workbook()
sheet(wb,"Sources","Source Register — The Forward-Looking Trader","Version 1.0 · 16 Nov 2026. Verification grades describe what was actually checked in this build session, not the reliability of the source in general.",
 ["ID","Source","Type","Date","URL / locator","Claims supported","Used in","Verification status","Method","Notes and limits"],[7,46,18,16,44,52,22,20,40,44],src,first=True)
claims=[
 ("C-01","Druckenmiller advises looking ~18–24 months ahead instead of at present conditions.","Investor statement","S1, S7, S5","Corroborated (secondary)","L1 §C1; Blueprint §1.1"),
 ("C-02","He uses a cyclical industry closing capacity as the example of why the future differs from the present.","Investor statement","S1, S7","Corroborated (secondary)","L1 §C1"),
 ("C-03","He describes most trades as evolving over ~18 months to 3 years; some ~1 year, some up to 5.","Investor statement","S2, S8","Corroborated (search excerpt of primary + secondary)","L1 §C1–C2; Blueprint §1.1"),
 ("C-04","He acts decisively and changes course quickly when facts change.","Investor statement","S2","Corroborated (publisher's description)","Blueprint §1.1, §9.1"),
 ("C-05","He considers contrarianism overrated.","Investor statement","S6","Corroborated (publisher's description)","L1 §F4"),
 ("C-06","He cautions against over-reliance on now-widespread technical signals (2026).","Investor statement, per handoff brief","S2","NOT corroborated this session","Not used in L1; flagged in Blueprint Wk 13"),
 ("C-07","Teva example: bought near $16, about $32 at recording, re-rated from ~6× to ~11.5–12× earnings.","Investor statement","S2, S8","Seen only in third-party auto-transcripts; verify before use","Reserved for Wk 9; not used in L1"),
 ("C-08","Early 2000: sold technology in January, bought back ~$6bn, lost ~$3bn in about six weeks.","Investor statement (own account)","S5","Corroborated (secondary transcripts)","L1 §F5"),
 ("C-09","Micron quarterly revenue, gross margin, net income and capex, FY2022–FY2026.","Fact — primary data","M1–M12","Verified (primary data)","L1 §B, D, E"),
 ("C-10","Micron management statements of 22 Dec 2022 on industry conditions, wafer starts, capex and headcount.","Fact — primary filing text","M2","Verified (filing text read)","L1 §B, D"),
 ("C-11","Micron share prices, market values and price returns on stated dates.","Fact — market data + arithmetic","M13","Verified (single feed)","L1 §D, E"),
 ("C-12","Trailing P/E ≈ 8.7× on 22 Dec 2022 and ≈ 10.6× on 12 Nov 2026.","Curriculum arithmetic","M1–M2, M11–M13","Computed; approximate","L1 §B, D, E"),
 ("C-13","AI-server demand contributed materially to the 2023–24 memory recovery and was not in the Dec 2022 filing.","Curriculum interpretation","M2 (absence from the Dec 2022 MD&A excerpt read); general knowledge","Partly verified: absence checked in the excerpt read; causal contribution is a widely held view, not sourced here","L1 §D, F3 — source properly before reuse"),
 ("C-14","Three clocks; why 18–24 months; Future-State Card; implied-probability framing; Calder Forge.","Curriculum inference / hypothetical","—","Original; labelled in text","L1 throughout"),
 ("C-15","Examples of industries where supply cuts did not restore profitability (coal, newspapers).","Curriculum illustration","—","Unsourced general illustration; to be replaced with dated cases in Wk 6","L1 §F1"),
]
sheet(wb,"Claims","Claim-level map","Each substantive claim, its evidence class and where it is used. 'Curriculum inference' = original to this course, not attributable to any investor.",
 ["Claim ID","Claim","Evidence class","Source IDs","Status (16 Nov 2026)","Where used"],[9,70,26,16,40,34],claims)
bib=[(a,b,c) for a,b,c in [
 ("Foundation","Jack D. Schwager, The New Market Wizards. HarperBusiness, 1992.","Contains a long interview with Druckenmiller."),
 ("Foundation","George Soros, The Alchemy of Finance. Simon & Schuster, 1987.","Reflexivity."),
 ("Foundation","Steven Drobny, Inside the House of Money. Wiley, 2006.","Macro practitioners."),
 ("Foresight / capital","Edward Chancellor (ed.), Capital Returns. Palgrave Macmillan, 2015.","See S3."),
 ("Foresight / capital","Philip E. Tetlock & Dan Gardner, Superforecasting. Crown, 2015.","Calibration, base rates."),
 ("Foresight / capital","Peter Schwartz, The Art of the Long View. Doubleday, 1991.","Scenario method."),
 ("Technology / strategy","Carlota Perez, Technological Revolutions and Financial Capital. Edward Elgar, 2002.","Diffusion."),
 ("Technology / strategy","Bruce Greenwald & Judd Kahn, Competition Demystified. Portfolio, 2005.","Competitive advantage."),
 ("Fundamentals","Michael J. Mauboussin & Alfred Rappaport, Expectations Investing. Columbia Business School Publishing, rev. ed. 2021.","See S4."),
 ("Fundamentals","Stephen H. Penman, Financial Statement Analysis and Security Valuation. McGraw-Hill (5th ed., 2013).","Accounting-based valuation."),
 ("Technical / process","John J. Murphy, Technical Analysis of the Financial Markets. New York Institute of Finance, 1999.",""),
 ("Technical / process","Stan Weinstein, Secrets for Profiting in Bull and Bear Markets. 1988.",""),
 ("Technical / process","Perry J. Kaufman, Trading Systems and Methods. Wiley (6th ed., 2019).",""),
 ("Technical / process","Brent Donnelly, Alpha Trader. 2021.",""),
 ("Psychology / risk","Mark Douglas, Trading in the Zone. 2000.",""),
 ("Psychology / risk","Jared Tendler, The Mental Game of Trading. 2021.",""),
 ("Psychology / risk","Nassim Nicholas Taleb, Fooled by Randomness. 2001.",""),
 ("Psychology / risk","Van K. Tharp, Trade Your Way to Financial Freedom. McGraw-Hill (2nd ed., 2007).",""),
 ("Portfolio / macro","Richard C. Grinold & Ronald N. Kahn, Active Portfolio Management. McGraw-Hill (2nd ed., 1999).",""),
 ("Portfolio / macro","Antti Ilmanen, Expected Returns. Wiley, 2011.",""),
 ("Portfolio / macro","Ralph Vince, The Mathematics of Money Management. Wiley, 1992.",""),
 ("Portfolio / macro","Edward O. Thorp, A Man for All Markets. Random House, 2017.",""),
 ("Portfolio / macro","Howard Marks, The Most Important Thing. Columbia University Press, 2011.",""),
]]
sheet(wb,"Bibliography","Bibliography","Publication details are standard references recorded from memory; none was re-checked against a publisher page in this session. Verify editions before citing page numbers.",
 ["Strand","Entry","Note"],[22,90,40],bib)
wb.save("../Source_Register.xlsx"); print("sources written")
