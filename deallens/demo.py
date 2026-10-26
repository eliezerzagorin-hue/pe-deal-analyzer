"""Entirely fictional case study for safe public demonstrations."""
from .memo import blank_memo

DEMO_COMPANY = "Northstar Facility Services (FICTIONAL DEMO)"

DEMO_HISTORY = [
    {"year": 2023, "revenue": 19.2, "ebitda": 2.7,
     "adjusted_ebitda": 3.0, "source_page": 0, "notes": "Fictional figures"},
    {"year": 2024, "revenue": 21.8, "ebitda": 3.2,
     "adjusted_ebitda": 3.6, "source_page": 0, "notes": "Fictional figures"},
    {"year": 2025, "revenue": 24.5, "ebitda": 3.9,
     "adjusted_ebitda": 4.3, "source_page": 0, "notes": "Fictional figures"},
]


def demo_memo() -> dict[str, str]:
    memo = blank_memo()
    memo.update({
        "management_employees": "FICTIONAL EXAMPLE. Founder-led management team supported by regional operators. Headcount and management retention are diligence items.",
        "products_services": "FICTIONAL EXAMPLE. Recurring facility maintenance and specialty cleaning services across commercial sites.",
        "general_operations": "FICTIONAL EXAMPLE. Multi-location branch structure with local dispatch and regional supervision. Verify labor utilization and route density.",
        "customer_base": "FICTIONAL EXAMPLE. Commercial and institutional customers under a mix of annual and multiyear service contracts. Obtain customer-level revenue detail.",
        "historical_financial_performance": "FICTIONAL EXAMPLE. FY2023–FY2025 revenue rises from $19.2M to $24.5M; reported EBITDA rises from $2.7M to $3.9M. Adjusted EBITDA includes unverified proposed add-backs.",
        "market_overview": "FICTIONAL EXAMPLE. Highly fragmented regional service market. Industry growth, local competitor shares, and pricing power require third-party research.",
        "growth_opportunities": "FICTIONAL EXAMPLE. Route density, cross-selling, selective price increases, and regional add-on acquisitions. No synergies assumed without validation.",
        "investment_merits_risks_mitigants": "FICTIONAL EXAMPLE. Merit: recurring contracts and diverse sites. Risks: wage inflation, customer concentration and integration. Mitigants to diligence: repricing clauses, contract duration and retention data.",
        "investment_thesis": "PRELIMINARY FICTIONAL THESIS. A potential regional buy-and-build platform if revenue recurrence, margins, customer retention, and acquisition pipeline validate.",
        "potential_exit_strategies_buyer_universe": "FICTIONAL HYPOTHESIS. Potential strategic facilities-services consolidators or sponsor-backed regional platforms; no actual buyer outreach or validation performed.",
        "recommendation": "PRELIMINARY FICTIONAL RECOMMENDATION. Advance only to focused diligence after confirming EBITDA adjustments, customer retention, wage pass-through, and quality of earnings.",
    })
    return memo
