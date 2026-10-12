"""Cross-language contract test.

The C# add-in serializes mail payloads with snake_case keys. This test pins
that contract so a rename on either side fails loudly instead of silently
dropping fields.
"""

from __future__ import annotations

from app.api import MailPayload
from app.indexing import IndexMailRequest

# Must match SemanticSearchClient.ToSnakeCaseDictionary in C#.
CSHARP_PAYLOAD_KEYS = {
    "store_id",
    "entry_id",
    "subject",
    "internet_message_id",
    "conversation_id",
    "sender_name",
    "sender_email",
    "recipients",
    "received_at",
    "sent_at",
    "last_modified_at",
    "folder_id",
    "folder_path",
    "attachment_names",
    "attachment_count",
    "body_text",
    "html_body",
    "owner_mailbox",
    "force_reembed",
}


def test_csharp_keys_accepted_by_api():
    api_fields = set(MailPayload.model_fields.keys())
    missing = CSHARP_PAYLOAD_KEYS - api_fields
    assert not missing, "C# sends keys the API does not accept: %s" % missing


def test_csharp_keys_accepted_by_index_request():
    request_fields = set(IndexMailRequest.__dataclass_fields__.keys())
    # force_reembed is consumed by the endpoint, not the dataclass.
    expected = CSHARP_PAYLOAD_KEYS - {"force_reembed"}
    missing = expected - request_fields
    assert not missing, (
        "C# sends keys IndexMailRequest does not accept: %s" % missing
    )


def test_no_pascal_case_keys():
    for key in CSHARP_PAYLOAD_KEYS:
        assert key == key.lower(), "%s must be snake_case" % key
        assert "_" in key or key in ("subject", "recipients"), (
            "%s should be snake_case" % key
        )
