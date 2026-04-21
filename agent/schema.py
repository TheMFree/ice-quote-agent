"""Quote extraction schema — single source of truth for fields the agent
extracts from incoming email content and injects into the proposal template.

Every field is optional. Missing values stay blank in the final document.
"""
from __future__ import annotations
from typing import List, Optional
from pydantic import BaseModel, Field


class PricingLine(BaseModel):
    description: Optional[str] = None
    amount: Optional[str] = None


class ScopeCategory(BaseModel):
    """A sub-section of Attachment A (Scope of Work Included)."""
    heading: str
    items: List[str] = Field(default_factory=list)


class QuoteData(BaseModel):
    project: Optional[str] = Field(None)
    agreement_no: Optional[str] = Field(None)
    wbs_no: Optional[str] = Field(None)
    client: Optional[str] = Field(None)
    owner_rep: Optional[str] = Field(
        None,
        description="Client-side project contact — the person at the "
                    "client company to whom this quote is addressed "
                    "(e.g. 'John Smith, Project Manager').",
    )
    location: Optional[str] = Field(None)
    proposal_date: Optional[str] = Field(None)
    quote_no: Optional[str] = None
    prepared_by: Optional[str] = Field(
        default="Michael G. Freeman, Jr. — President",
    )
    valid_for: Optional[str] = Field(default="30 Days from Proposal Date")

    document_type: Optional[str] = Field(default="LUMP SUM PROPOSAL")

    scope_intro: Optional[str] = Field(None)
    scope_bullets: List[str] = Field(default_factory=list)

    material_amount: Optional[str] = None
    labor_equipment_amount: Optional[str] = None
    total_amount: Optional[str] = None
    additional_pricing_lines: List[PricingLine] = Field(default_factory=list)

    total_man_hours: Optional[str] = None
    crew_size: Optional[str] = None
    estimated_duration: Optional[str] = None
    work_schedule: Optional[str] = None

    scope_categories: List[ScopeCategory] = Field(default_factory=list)

    long_lead_items: List[str] = Field(default_factory=list)
    assumptions: List[str] = Field(default_factory=list)
    pending_clarifications: List[str] = Field(default_factory=list)

    additional_exclusions: List[str] = Field(default_factory=list)

    extraction_notes: Optional[str] = Field(None)
