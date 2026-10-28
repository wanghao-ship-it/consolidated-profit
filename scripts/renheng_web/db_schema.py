"""PostgreSQL业务表结构。Checkpoint表由LangGraph独立管理。"""
from __future__ import annotations

from sqlalchemy import (
    BigInteger,
    Boolean,
    CheckConstraint,
    Column,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    JSON,
    MetaData,
    String,
    Table,
    Text,
    UniqueConstraint,
)

metadata = MetaData()

users = Table(
    "users",
    metadata,
    Column("user_id", String(36), primary_key=True),
    Column("username", String(80), nullable=False, unique=True),
    Column("display_name", String(120), nullable=False),
    Column("role", String(24), nullable=False),
    Column("password_hash", String(256), nullable=False),
    Column("is_active", Boolean, nullable=False, default=True),
    Column("created_at", DateTime(timezone=True), nullable=False),
    CheckConstraint("role in ('FINANCE', 'REVIEWER', 'ADMIN')", name="ck_users_role"),
)

workflow_runs = Table(
    "workflow_runs",
    metadata,
    Column("run_id", String(48), primary_key=True),
    Column("thread_id", String(255), nullable=False, unique=True),
    Column("year", Integer, nullable=False),
    Column("month", Integer, nullable=False),
    Column("created_by", String(36), ForeignKey("users.user_id"), nullable=False),
    Column("status", String(32), nullable=False),
    Column("preflight_status", String(24), nullable=False, default="NOT_RUN"),
    Column("preflight_result", JSON),
    Column("formal_report_object_key", Text),
    Column("error_message", Text),
    Column("created_at", DateTime(timezone=True), nullable=False),
    Column("updated_at", DateTime(timezone=True), nullable=False),
    Column("completed_at", DateTime(timezone=True)),
    CheckConstraint("month between 1 and 12", name="ck_workflow_runs_month"),
)
Index("ix_workflow_runs_created_at", workflow_runs.c.created_at)
Index("ix_workflow_runs_status", workflow_runs.c.status)

workflow_files = Table(
    "workflow_files",
    metadata,
    Column("file_id", String(36), primary_key=True),
    Column("run_id", String(48), ForeignKey("workflow_runs.run_id", ondelete="CASCADE"), nullable=False),
    Column("role", String(24), nullable=False),
    Column("artifact_key", String(160), nullable=False),
    Column("object_key", Text, nullable=False, unique=True),
    Column("original_filename", Text, nullable=False),
    Column("local_path", Text, nullable=False),
    Column("content_type", String(255), nullable=False),
    Column("size_bytes", BigInteger, nullable=False),
    Column("sha256", String(64), nullable=False),
    Column("version", Integer, nullable=False),
    Column("is_current", Boolean, nullable=False, default=True),
    Column("created_by", String(36), ForeignKey("users.user_id"), nullable=False),
    Column("created_at", DateTime(timezone=True), nullable=False),
    UniqueConstraint("run_id", "role", "artifact_key", "version", name="uq_workflow_file_version"),
)
Index(
    "uq_workflow_files_current",
    workflow_files.c.run_id,
    workflow_files.c.role,
    workflow_files.c.artifact_key,
    unique=True,
    postgresql_where=workflow_files.c.is_current.is_(True),
    sqlite_where=workflow_files.c.is_current.is_(True),
)
Index("ix_workflow_files_run", workflow_files.c.run_id, workflow_files.c.created_at)

approval_records = Table(
    "approval_records",
    metadata,
    Column("approval_id", String(36), primary_key=True),
    Column("run_id", String(48), ForeignKey("workflow_runs.run_id", ondelete="CASCADE"), nullable=False),
    Column("review_round", Integer, nullable=False),
    Column("reviewer_id", String(36), ForeignKey("users.user_id"), nullable=False),
    Column("reason", Text, nullable=False),
    Column("corrected_revenue", JSON, nullable=False),
    Column("before_differences", JSON, nullable=False),
    Column("created_at", DateTime(timezone=True), nullable=False),
    UniqueConstraint("run_id", "review_round", name="uq_approval_run_round"),
)
Index("ix_approval_records_run", approval_records.c.run_id, approval_records.c.created_at)

run_events = Table(
    "run_events",
    metadata,
    Column("event_id", String(36), primary_key=True),
    Column("run_id", String(48), ForeignKey("workflow_runs.run_id", ondelete="CASCADE"), nullable=False),
    Column("actor_id", String(36), ForeignKey("users.user_id")),
    Column("event_type", String(48), nullable=False),
    Column("details", JSON, nullable=False),
    Column("created_at", DateTime(timezone=True), nullable=False),
)
Index("ix_run_events_run", run_events.c.run_id, run_events.c.created_at)
