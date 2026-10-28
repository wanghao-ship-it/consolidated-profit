"""Create business, file-index, user, and review tables."""
import sqlalchemy as sa
from alembic import op

revision = "0001_initial"
down_revision = None
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "users",
        sa.Column("user_id", sa.String(36), primary_key=True),
        sa.Column("username", sa.String(80), nullable=False, unique=True),
        sa.Column("display_name", sa.String(120), nullable=False),
        sa.Column("role", sa.String(24), nullable=False),
        sa.Column("password_hash", sa.String(256), nullable=False),
        sa.Column("is_active", sa.Boolean(), nullable=False, server_default=sa.true()),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.CheckConstraint("role in ('FINANCE', 'REVIEWER', 'ADMIN')", name="ck_users_role"),
    )
    op.create_table(
        "workflow_runs",
        sa.Column("run_id", sa.String(48), primary_key=True),
        sa.Column("thread_id", sa.String(255), nullable=False, unique=True),
        sa.Column("year", sa.Integer(), nullable=False),
        sa.Column("month", sa.Integer(), nullable=False),
        sa.Column("created_by", sa.String(36), sa.ForeignKey("users.user_id"), nullable=False),
        sa.Column("status", sa.String(32), nullable=False),
        sa.Column("preflight_status", sa.String(24), nullable=False, server_default="NOT_RUN"),
        sa.Column("preflight_result", sa.JSON()),
        sa.Column("formal_report_object_key", sa.Text()),
        sa.Column("error_message", sa.Text()),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("completed_at", sa.DateTime(timezone=True)),
        sa.CheckConstraint("month between 1 and 12", name="ck_workflow_runs_month"),
    )
    op.create_index("ix_workflow_runs_created_at", "workflow_runs", ["created_at"])
    op.create_index("ix_workflow_runs_status", "workflow_runs", ["status"])
    op.create_table(
        "workflow_files",
        sa.Column("file_id", sa.String(36), primary_key=True),
        sa.Column("run_id", sa.String(48), sa.ForeignKey("workflow_runs.run_id", ondelete="CASCADE"), nullable=False),
        sa.Column("role", sa.String(24), nullable=False),
        sa.Column("artifact_key", sa.String(160), nullable=False),
        sa.Column("object_key", sa.Text(), nullable=False, unique=True),
        sa.Column("original_filename", sa.Text(), nullable=False),
        sa.Column("local_path", sa.Text(), nullable=False),
        sa.Column("content_type", sa.String(255), nullable=False),
        sa.Column("size_bytes", sa.BigInteger(), nullable=False),
        sa.Column("sha256", sa.String(64), nullable=False),
        sa.Column("version", sa.Integer(), nullable=False),
        sa.Column("is_current", sa.Boolean(), nullable=False, server_default=sa.true()),
        sa.Column("created_by", sa.String(36), sa.ForeignKey("users.user_id"), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.UniqueConstraint("run_id", "role", "artifact_key", "version", name="uq_workflow_file_version"),
    )
    op.create_index(
        "uq_workflow_files_current",
        "workflow_files",
        ["run_id", "role", "artifact_key"],
        unique=True,
        postgresql_where=sa.text("is_current IS TRUE"),
        sqlite_where=sa.text("is_current IS TRUE"),
    )
    op.create_index("ix_workflow_files_run", "workflow_files", ["run_id", "created_at"])
    op.create_table(
        "approval_records",
        sa.Column("approval_id", sa.String(36), primary_key=True),
        sa.Column("run_id", sa.String(48), sa.ForeignKey("workflow_runs.run_id", ondelete="CASCADE"), nullable=False),
        sa.Column("review_round", sa.Integer(), nullable=False),
        sa.Column("reviewer_id", sa.String(36), sa.ForeignKey("users.user_id"), nullable=False),
        sa.Column("reason", sa.Text(), nullable=False),
        sa.Column("corrected_revenue", sa.JSON(), nullable=False),
        sa.Column("before_differences", sa.JSON(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.UniqueConstraint("run_id", "review_round", name="uq_approval_run_round"),
    )
    op.create_index("ix_approval_records_run", "approval_records", ["run_id", "created_at"])
    op.create_table(
        "run_events",
        sa.Column("event_id", sa.String(36), primary_key=True),
        sa.Column("run_id", sa.String(48), sa.ForeignKey("workflow_runs.run_id", ondelete="CASCADE"), nullable=False),
        sa.Column("actor_id", sa.String(36), sa.ForeignKey("users.user_id")),
        sa.Column("event_type", sa.String(48), nullable=False),
        sa.Column("details", sa.JSON(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    )
    op.create_index("ix_run_events_run", "run_events", ["run_id", "created_at"])


def downgrade() -> None:
    op.drop_index("ix_run_events_run", table_name="run_events")
    op.drop_table("run_events")
    op.drop_index("ix_approval_records_run", table_name="approval_records")
    op.drop_table("approval_records")
    op.drop_index("ix_workflow_files_run", table_name="workflow_files")
    op.drop_index("uq_workflow_files_current", table_name="workflow_files")
    op.drop_table("workflow_files")
    op.drop_index("ix_workflow_runs_status", table_name="workflow_runs")
    op.drop_index("ix_workflow_runs_created_at", table_name="workflow_runs")
    op.drop_table("workflow_runs")
    op.drop_table("users")
