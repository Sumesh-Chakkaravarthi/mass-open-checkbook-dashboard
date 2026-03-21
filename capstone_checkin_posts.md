# Capstone Check-in: Module 6 Discussion Posts

## Primary Prompt: Dashboard Summary

![Dashboard Screenshot](/Users/sumesh/Projects/Antigravity/Capstone/dashboard_screenshot.png)

**Dashboard Summary:**

For my individual capstone project, I am developing an interactive Plotly Dash analytical dashboard for Summit Global Consulting. The primary business problem this project addresses is the lengthy vendor onboarding procedure and market entry delays caused by fragmented, non-standardized insights into vendor operations and their Supplier Diversity Office (SDO) commitments. 

The dashboard serves as a unified decision-making system built upon the Massachusetts Open Checkbook vendor contract data. It is specifically designed for non-technical business professionals—such as business development managers, procurement analysts, and executive leadership—to help them easily locate, evaluate, and rank high-value IT-industry vendors. 

Structurally, the dashboard features five overarching KPI cards (Total Vendors, Average SDO Commitment, Procurement Categories, IT Vendors, and Industries) that give executives an immediate macroscopic view of the data. It is divided into four distinct interactive tabs: "IT Sector SDO", "Cross-Category", "Vendor Coverage", and "Industry Analysis". Users can dynamically filter data by IT sub-categories to see which companies lead in SDO commitments, view treemaps of vendor distributions, and compare national versus local business presence. 

Instead of relying on opaque predictive models, the dashboard utilizes a weighted scoring framework (e.g., prioritizing SDO Commitment and Addressable Spend) to rank operational accessibility and financial viability. This conscious design choice prioritizes operational intelligence and strategic clarity over simple algorithmic accuracy, aligning with Obermeyer and Emanuel's (2016) assertion that analytical model systems need proper controls and domain understanding prior to implementation. By transforming raw, disparate data into structured vendor intelligence, this analytical platform successfully aligns organizational purchasing decisions with both operational efficiency and equitable, sustainable business practices, grounded firmly in strong data governance principles (Khatri & Brown, 2010).

**References:**

Khatri, V., & Brown, C. V. (2010). Designing data governance. *Communications of the ACM, 53*(1), 148–152. https://doi.org/10.1145/1629175.1629210

Obermeyer, Z., & Emanuel, E. J. (2016). Predicting the future—Big data, machine learning, and clinical medicine. *The New England Journal of Medicine, 375*(13), 1216–1219. https://www.nejm.org/doi/10.1056/NEJMp1606181

---

## Secondary Prompt: Peer Replies

### Reply 1: To Prithvi Yadav Pulicherla
**Word Count: 147 words**

Hi Prithvi, 

Great job on your Power BI dashboard! I really appreciate how you structured the visual layout, specifically the way you used the sector concentration chart to highlight Healthcare & Biotechnology and Finance & Insurance as key focus areas. It provides immediate value and makes the data finding highly accessible for non-technical stakeholders looking to prioritize strategic outreach. This aligns well with Davenport and Harris (2007), who emphasize that the actionable value of analytics is only realized when it matches the operational requirements of the leadership team. 

As an area for improvement, have you considered adding a specific metric that defines data completeness, such as contact verification, to your upcoming enhancements? According to Khatri and Brown (2010), data trustworthiness should take priority over data completeness in governance decisions. Explicitly filtering out unverified contacts could strengthen the defensibility of the decisions leadership makes using your tool.

**References:**
- Davenport, T. H., & Harris, J. G. (2007). *Competing on analytics: The new science of winning*. Harvard Business School Press.
- Khatri, V., & Brown, C. V. (2010). Designing data governance. *Communications of the ACM, 53*(1), 148–152.

### Reply 2: To Krishna Murari Sharma
**Word Count: 143 words**

Hi Krishna, 

Excellent work on your conceptual model and your dashboard presentation. The way you prioritized vendor filtering into actionable tiers—specifically calling out the 48 high-priority vendors versus the 38 flagged for de-prioritization—provides a massive amount of strategic clarity to the end-user. Utilizing the risk score radar to highlight low activity levels is also a very clean way to display multi-dimensional metrics.

I do have a clarification question regarding your risk framework: when evaluating these different vendors, how does your system handle missing SDO data (since only 89 of 212 have it) to prevent scoring bias in the radar chart? Obermeyer and Emanuel (2016) caution that analytical models need proper controls and must remain unbiased to ensure ethical implementation. Adding a visual indicator or transparency note regarding how missing SDO values affect the overall risk score would be a great enhancement!

**References:**
- Obermeyer, Z., & Emanuel, E. J. (2016). Predicting the future—Big data, machine learning, and clinical medicine. *The New England Journal of Medicine, 375*(13), 1216–1219.
