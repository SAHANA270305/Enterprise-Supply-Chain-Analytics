Enterprise Demand Forecasting: Cisco Supply Chain (CFL Phase 2)
📌 Project Overview
This repository contains the demand forecasting models and analytical framework developed for the Cisco Supply Chain Forecasting (CFL) Competition - Phase 2.

The objective of this project was to forecast the FY26 Q2 demand for Cisco's top 20 cost-ranked hardware products (PLIDs). Rather than relying on a universal algorithm that treats all variance as true demand, this project utilizes a behavior-driven demand planning strategy to protect working capital and prevent the over-purchasing of high-cost infrastructure assets.

Final Result: A structurally sound, financially de-risked forecast of 63,486 units (~3.6% QoQ growth), achieving a balance between robust supply for high-volume endpoints and strict capital protection on volatile, high-cost SKUs.

⚠️ The Business Problem
A single time-series forecasting model cannot fit an entire enterprise hardware portfolio. Cisco’s product lines exhibit a strict Volume-to-Value Inversion:

High-volume, low-cost endpoints (e.g., IP Phones) operate on massive, stable demand (>11,000 units/quarter).

Low-volume, high-cost infrastructure (e.g., 400G Data Center Spines) operate on highly volatile, project-based demand (<150 units/quarter).

Historical data is often polluted by one-off events (bulk orders, project rollouts). Feeding unadjusted statistical averages into a supply chain model exposes the business to millions of dollars in inventory liability.

🧠 Forecasting Methodology
To solve for this portfolio variance, we engineered a three-step segmented forecasting architecture tailored to the specific demand physics of each product.

Phase 1: Demand Segmentation (Syntetos-Boylan Classification)
We calculated two core metrics from historical actuals to classify the 20 products into distinct behavioral profiles:

ADI (Average Demand Interval): Measures the regularity of demand in time.

CV² (Coefficient of Variation squared): Measures the variance in demand quantity.

Using these metrics, the portfolio was segmented into four categories:

Smooth: Steady and regular demand.

Erratic: Regular timing but highly variable quantities.

Lumpy: Sudden, large bursts (project-driven) with periods of zero demand.

Intermittent: Rare and irregular demand.

Phase 2: Tailored Modeling
Once segmented, we applied specific forecasting logic to each bucket rather than a "one-size-fits-all" approach:

Smooth: Blended an Ensemble Consensus Model (combining Planner, Marketing, and Data Science inputs) with recent run-rates.

Lumpy: Weighted historical "Big Deal" pipeline data and VMS (Vertical Market Segments) trends to capture project-driven bursts while ignoring baseline noise.

Erratic: Utilized a Seasonal Time Series blend to capture historical ups and downs while heavily smoothing peak anomalies.

Intermittent: Relied on Bottom-Up SCMS Channel Aggregation to filter out top-level noise.

Phase 3: FVA Governance Layer (Forecast Value Added)
To prevent algorithmic over-forecast bias, we implemented a strict "Human-in-the-Loop" governance layer, mandating visual historical trend checks for the Top 5 cost-weighted products.

Action Taken: The unadjusted statistical model suggested an aggressive 4,544 units for Rank 3 (ROUTER Branch LTE) based on historical spikes. Recognizing these spikes as episodic noise rather than structural demand, we applied a manual FVA override, capping the forecast at 4,200 units.

Impact: This single manual intervention removed ~344 excess units of high-cost liability from the system, protecting working capital without compromising Q2 service levels.

📊 Key Findings & Impact
Financial Risk Reduction: Successfully removed algorithmic over-forecast bias in top-tier SKUs, optimizing working capital utilization.

Model Robustness: Closed the gap between validation and backtest performance by eliminating data leakage and utilizing segment-driven signals rather than pure historical averaging.

Touchless Forecasting: By trusting the algorithm for lower-cost "Smooth" items (C-class inventory) and applying human FVA only to Top 5 "Lumpy/Erratic" items (A-class inventory), the framework scales efficiently for enterprise production.

🛠️ Tools & Technologies Used
Python (Pandas, NumPy): Data cleaning, Syntetos-Boylan calculations (ADI/CV²), and segmentation logic.

Statistical Modeling: Time-series smoothing, ensemble consensus averaging.

Excel: Final model aggregation and FVA override tracking.

Supply Chain Principles: ABC Analysis, S&OP Governance, FVA (Forecast Value Added).
