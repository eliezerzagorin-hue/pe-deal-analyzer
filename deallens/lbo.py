"""Deterministic five-year, single-tranche LBO underwriting engine."""
from dataclasses import dataclass
from typing import Any


@dataclass
class Assumptions:
    entry_revenue: float = 24.5  # $ millions
    entry_ebitda: float = 3.9
    entry_multiple: float = 7.0
    debt_multiple: float = 3.0
    deal_fees_pct: float = 0.02
    revenue_growth: float = 0.05
    entry_margin: float = 0.15918367347
    annual_margin_change: float = 0.005
    interest_rate: float = 0.09
    tax_rate: float = 0.25
    da_pct_revenue: float = 0.02
    capex_pct_revenue: float = 0.03
    nwc_pct_revenue: float = 0.10
    exit_multiple: float = 7.0
    hold_years: int = 5


def _irr(cash_flows: list[float]) -> float | None:
    """Periodic IRR with no third-party finance calculation dependency."""
    if not any(x < 0 for x in cash_flows) or not any(x > 0 for x in cash_flows):
        return None

    def npv(rate: float) -> float:
        return sum(cf / ((1 + rate) ** i) for i, cf in enumerate(cash_flows))

    lo, hi = -0.9999, 1.0
    while npv(lo) * npv(hi) > 0 and hi < 1e5:
        hi *= 2
    if npv(lo) * npv(hi) > 0:
        return None
    for _ in range(160):
        mid = (lo + hi) / 2
        val = npv(mid)
        if abs(val) < 1e-10:
            return mid
        if npv(lo) * val <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def calculate_lbo(a: Assumptions) -> dict[str, Any]:
    """Returns annual projections and cash-flow-based returns; all amounts in $m.

    Caveats: assumes debt-free / cash-free purchase price, fees funded by sponsor,
    a single debt tranche, interest on opening debt, no distributions until exit,
    and sponsor top-ups for cash deficits.
    """
    if a.entry_revenue <= 0 or a.entry_ebitda <= 0:
        raise ValueError("Entry revenue and EBITDA must be positive.")
    if a.hold_years < 1 or a.hold_years > 15:
        raise ValueError("Hold period must be between one and fifteen years.")
    if a.entry_multiple <= 0 or a.exit_multiple <= 0:
        raise ValueError("Entry and exit multiples must be positive.")
    if a.debt_multiple < 0 or a.debt_multiple > a.entry_multiple:
        raise ValueError("Debt multiple cannot exceed entry enterprise-value multiple.")
    if not (0 <= a.tax_rate <= 1 and 0 <= a.interest_rate <= 1):
        raise ValueError("Interest rate and tax rate must be between 0% and 100%.")
    if a.revenue_growth <= -1:
        raise ValueError("Annual revenue decline must be less than 100%.")
    if a.entry_margin <= 0 or a.entry_margin > 1:
        raise ValueError("Entry EBITDA margin must be positive and no higher than 100%.")
    if any(x < 0 or x > 1 for x in
           (a.deal_fees_pct, a.da_pct_revenue, a.capex_pct_revenue, a.nwc_pct_revenue)):
        raise ValueError("Cost and working-capital ratios must be within 0% and 100%.")

    enterprise_value = a.entry_ebitda * a.entry_multiple
    opening_debt = a.entry_ebitda * a.debt_multiple
    fees = enterprise_value * a.deal_fees_pct
    initial_equity = enterprise_value + fees - opening_debt
    if initial_equity <= 0:
        raise ValueError("Entry sponsor equity must be positive.")

    rows: list[dict[str, float | int]] = []
    revenue = a.entry_revenue
    outstanding_debt = opening_debt
    investor_cash_flows = [-initial_equity]
    for year in range(1, a.hold_years + 1):
        previous_revenue = revenue
        revenue = revenue * (1 + a.revenue_growth)
        margin = a.entry_margin + a.annual_margin_change * year
        if not 0 < margin <= 1:
            raise ValueError("Projected EBITDA margin falls outside 0% to 100%.")
        ebitda = revenue * margin
        depreciation = revenue * a.da_pct_revenue
        ebit = ebitda - depreciation
        interest = outstanding_debt * a.interest_rate
        taxable_income = max(ebit - interest, 0)
        cash_taxes = taxable_income * a.tax_rate
        capex = revenue * a.capex_pct_revenue
        nwc_investment = (revenue - previous_revenue) * a.nwc_pct_revenue
        free_cash_flow = ebitda - interest - cash_taxes - capex - nwc_investment
        debt_repayment = min(outstanding_debt, max(0.0, free_cash_flow))
        equity_injection = max(0.0, -free_cash_flow)
        outstanding_debt -= debt_repayment
        rows.append({
            "year": year, "revenue": revenue, "ebitda_margin": margin,
            "ebitda": ebitda, "da": depreciation, "ebit": ebit,
            "interest": interest, "cash_taxes": cash_taxes,
            "capex": capex, "nwc_investment": nwc_investment,
            "free_cash_flow": free_cash_flow, "debt_repayment": debt_repayment,
            "equity_injection": equity_injection, "ending_debt": outstanding_debt,
        })
        investor_cash_flows.append(-equity_injection)

    exit_ebitda = rows[-1]["ebitda"]
    exit_ev = exit_ebitda * a.exit_multiple
    exit_equity = max(0.0, exit_ev - outstanding_debt)
    investor_cash_flows[-1] += exit_equity
    paid_in = initial_equity + sum(r["equity_injection"] for r in rows)
    moic = exit_equity / paid_in
    return {
        "entry_ev": enterprise_value, "entry_debt": opening_debt,
        "deal_fees": fees, "initial_equity": initial_equity,
        "total_paid_in": paid_in, "projection": rows, "exit_ev": exit_ev,
        "exit_ebitda": exit_ebitda, "exit_debt": outstanding_debt,
        "exit_equity": exit_equity, "moic": moic,
        "irr": _irr(investor_cash_flows),
        "investor_cash_flows": investor_cash_flows,
        "additional_equity": paid_in - initial_equity,
    }
