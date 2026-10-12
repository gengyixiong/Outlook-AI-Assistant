"""Retrieval quality evaluation on a fixed simulated mailbox.

This is a *simulated* evaluation: the fake embedding model is concept-based,
so it validates the retrieval pipeline and fusion mechanics, not the real
EmbeddingGemma 2 quality. Real-mailbox numbers must be collected separately
with the actual model; see docs/ARCHITECTURE.md.
"""

from __future__ import annotations

import math

import pytest

from app.indexing import Indexer
from app.search import HybridSearch

from conftest import FakeEmbeddingModel, make_mail

# (query, expected_entry_id, sender)
DATASET = [
    ("之前印度客户问过幕墙玻璃 Panel 能不能单独编辑的邮件", "E1",
     "india@example.com"),
    ("找一下之前那个代理商问 IFC4 认证的邮件", "E2", "reseller@example.com"),
    ("上个月有人反馈 Revit Link 导入后材质丢失", "E3", "user@example.com"),
    ("之前印度客户询问从 RVT Link 选择指定视图的问题", "E4",
     "india@example.com"),
    ("有个客户反馈 GstarBIM 占用了 24GB 内存", "E5", "customer@example.com"),
    ("Richard 之前发过一个关于渠道培训的安排", "E6", "richard@example.com"),
    ("找一下带有 Excel 附件的产品版本对比邮件", "E7", "pm@example.com"),
]

MAILS = [
    dict(entry_id="E1", internet_message_id="m1",
         sender_email="india@example.com", sender_name="India Client",
         subject="Re: Curtain wall edit",
         body_text="Regarding the curtain wall edit, can we also edit the "
                   "Panels individually?"),
    dict(entry_id="E2", internet_message_id="m2",
         sender_email="reseller@example.com", sender_name="Reseller",
         subject="IFC4 certification",
         body_text="We need the IFC4 certification for the reseller program."),
    dict(entry_id="E3", internet_message_id="m3",
         sender_email="user@example.com", sender_name="User",
         subject="Material loss after import",
         body_text="After importing the Revit Link the material is lost."),
    dict(entry_id="E4", internet_message_id="m4",
         sender_email="india@example.com", sender_name="India Client",
         subject="Select view from link",
         body_text="How do we select a specific view from the RVT Link?"),
    dict(entry_id="E5", internet_message_id="m5",
         sender_email="customer@example.com", sender_name="Customer",
         subject="Memory usage",
         body_text="GstarBIM used 24GB of memory on the client machine."),
    dict(entry_id="E6", internet_message_id="m6",
         sender_email="richard@example.com", sender_name="Richard",
         subject="Channel training plan",
         body_text="Here is the channel training schedule for next quarter."),
    dict(entry_id="E7", internet_message_id="m7",
         sender_email="pm@example.com", sender_name="PM",
         subject="Product version comparison",
         body_text="Attached is the product version comparison spreadsheet.",
         attachment_names="version-compare.xlsx", attachment_count=1),
    # Distractor mails that must not outrank the targets.
    dict(entry_id="D1", internet_message_id="d1",
         sender_email="noise@example.com", sender_name="Noise",
         subject="Office lunch",
         body_text="Lunch is at noon in the canteen."),
    dict(entry_id="D2", internet_message_id="d2",
         sender_email="noise@example.com", sender_name="Noise",
         subject="Holiday schedule",
         body_text="The holiday schedule is attached."),
]


@pytest.fixture()
def evaluated(settings, database):
    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    for mail in MAILS:
        indexer.index_mail(make_mail(**mail))
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    results = []
    for query, expected, sender in DATASET:
        response = searcher.search(
            query, senders=[sender] if sender else None, top_k=10
        )
        results.append((query, expected, response.results))
    return results


def test_recall_at_5(evaluated):
    hits = 0
    for _, expected, results in evaluated:
        top = [r.entry_id for r in results[:5]]
        if expected in top:
            hits += 1
    recall = hits / len(evaluated)
    print("Recall@5 = %.3f" % recall)
    assert recall >= 0.85


def test_hit_at_1(evaluated):
    hits = sum(
        1 for _, expected, results in evaluated
        if results and results[0].entry_id == expected
    )
    hit1 = hits / len(evaluated)
    print("Hit@1 = %.3f" % hit1)
    assert hit1 >= 0.7


def test_mrr_at_10(evaluated):
    total = 0.0
    for _, expected, results in evaluated:
        reciprocal = 0.0
        for rank, result in enumerate(results[:10], start=1):
            if result.entry_id == expected:
                reciprocal = 1.0 / rank
                break
        total += reciprocal
    mrr = total / len(evaluated)
    print("MRR@10 = %.3f" % mrr)
    assert mrr >= 0.75


def test_every_query_returns_real_indexed_mail(evaluated):
    """No fabricated results: every hit must be a real indexed entry."""

    real_ids = {mail["entry_id"] for mail in MAILS}
    for _, _, results in evaluated:
        for result in results:
            assert result.entry_id in real_ids
            assert result.subject
