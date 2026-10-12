"""Local semantic email search service for Outlook AI Assistant.

The service owns the embedding model and all SQLite storage. The C# Outlook
add-in never touches the database directly and never calls Outlook COM from
this process.
"""

__version__ = "0.5.0"
