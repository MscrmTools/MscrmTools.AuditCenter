using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace MsCrmTools.AuditCenter.Forms
{
    public partial class AuditCleanupForm : Form
    {
        private readonly IOrganizationService service;
        private readonly List<EntityMetadata> auditEnabledEntities;

        public AuditCleanupForm(IOrganizationService service, IEnumerable<EntityMetadata> entities)
        {
            InitializeComponent();

            this.service = service;
            auditEnabledEntities = (entities ?? new List<EntityMetadata>())
                .Where(e => e.IsAuditEnabled != null && e.IsAuditEnabled.Value)
                .OrderBy(e => e.LogicalName)
                .ToList();

            dtpFrom.Value = DateTime.UtcNow.Date.AddMonths(-1);
            dtpTo.Value = DateTime.UtcNow.Date.AddDays(-1);

            txtFilter.TextChanged += (s, e) => PopulateTables();

            PopulateTables();
        }

        private void PopulateTables()
        {
            var filter = txtFilter.Text.Trim();

            var previouslyChecked = new HashSet<string>(
                clbTables.CheckedItems.Cast<TableItem>().Select(i => i.LogicalName));

            clbTables.BeginUpdate();
            clbTables.Items.Clear();

            foreach (var emd in auditEnabledEntities)
            {
                var displayName = emd.DisplayName?.UserLocalizedLabel?.Label ?? emd.LogicalName;

                if (filter.Length > 0
                    && emd.LogicalName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0
                    && displayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                var index = clbTables.Items.Add(new TableItem(emd.LogicalName, displayName, emd.ObjectTypeCode));
                if (previouslyChecked.Contains(emd.LogicalName))
                {
                    clbTables.SetItemChecked(index, true);
                }
            }

            clbTables.EndUpdate();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            var selectedTables = clbTables.CheckedItems.Cast<TableItem>().ToList();

            if (selectedTables.Count == 0)
            {
                MessageBox.Show(this, "Select at least one table.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DateTime fromUtc, toUtc;
            if (!TryGetUtcRange(out fromUtc, out toUtc))
            {
                return;
            }

            if (!chkConfirm.Checked)
            {
                MessageBox.Show(this, "You must confirm that you understand this operation is irreversible.",
                    "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var summary = string.Format(
                "You are about to permanently delete audit records for {0} table(s):\r\n\r\n" +
                "  {1}\r\n\r\n" +
                "Changed Date from {2:yyyy-MM-dd} 00:00 UTC (included)\r\n" +
                "               to {3:yyyy-MM-dd} 00:00 UTC (excluded)\r\n\r\n" +
                "This cannot be undone. Continue?",
                selectedTables.Count,
                string.Join(", ", selectedTables.Select(x => x.LogicalName)),
                fromUtc, toUtc);

            if (MessageBox.Show(this, summary, "Confirm bulk delete",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            SubmitBulkDelete(selectedTables, fromUtc, toUtc);
        }

        private void SubmitBulkDelete(List<TableItem> selectedTables, DateTime fromUtc, DateTime toUtc)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                btnSubmit.Enabled = false;

                var querySet = new List<QueryExpression>();

                foreach (var table in selectedTables)
                {
                    querySet.Add(BuildAuditQuery(table, fromUtc, toUtc, new ColumnSet(false)));
                }

                var request = new BulkDeleteRequest
                {
                    JobName = string.Format("Audit Center cleanup - {0} - Changed Date {1:yyyy-MM-dd} to {2:yyyy-MM-dd} UTC",
                        string.Join(",", selectedTables.Select(x => x.LogicalName)),
                        fromUtc, toUtc.AddDays(-1)),
                    QuerySet = querySet.ToArray(),
                    StartDateTime = DateTime.Now,
                    RecurrencePattern = string.Empty,
                    SendEmailNotification = false,
                    ToRecipients = new Guid[0],
                    CCRecipients = new Guid[0]
                };

                var response = (BulkDeleteResponse)service.Execute(request);

                lblResult.ForeColor = System.Drawing.Color.DarkGreen;
                lblResult.Text = string.Format(
                    "Bulk delete job submitted. Job Id: {0}\r\nThe job runs asynchronously; monitor it under System Jobs.",
                    response.JobId);
            }
            catch (Exception ex)
            {
                lblResult.ForeColor = System.Drawing.Color.DarkRed;
                lblResult.Text = "Failed to submit bulk delete job: " + ex.Message;
                MessageBox.Show(this, "An error occured: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                btnSubmit.Enabled = true;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Builds the UTC range from the pickers. Only the date part is used: the range always
        /// starts at 00:00:00 of the "From" day and ends at 00:00:00 of the day AFTER the "To" day.
        /// The upper bound is therefore exclusive, which makes the "To" day fully included without
        /// relying on fractional seconds (Dataverse truncates audit createdon to whole seconds and
        /// the Bulk Delete UI cannot render sub-second literals).
        /// The values are built with DateTimeKind.Utc so the SDK performs no local-to-UTC shift.
        /// </summary>
        private bool TryGetUtcRange(out DateTime fromUtc, out DateTime toUtcExclusive)
        {
            var f = dtpFrom.Value.Date;
            var t = dtpTo.Value.Date;

            fromUtc = new DateTime(f.Year, f.Month, f.Day, 0, 0, 0, DateTimeKind.Utc);
            toUtcExclusive = new DateTime(t.Year, t.Month, t.Day, 0, 0, 0, DateTimeKind.Utc).AddDays(1);

            if (fromUtc >= toUtcExclusive)
            {
                MessageBox.Show(this, "The 'From' day must not be later than the 'To' day.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private QueryExpression BuildAuditQuery(TableItem table, DateTime fromUtc, DateTime toUtcExclusive, ColumnSet columns)
        {
            var query = new QueryExpression("audit")
            {
                ColumnSet = columns,
                Criteria = new FilterExpression(LogicalOperator.And)
            };

            // The audit.objecttypecode attribute is an EntityName-typed picklist. Passing the
            // numeric object type code makes the filter resolvable by the Bulk Delete system job
            // UI, which otherwise renders the Entity condition with an empty value.
            if (table.ObjectTypeCode.HasValue)
            {
                query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, table.ObjectTypeCode.Value);
            }
            else
            {
                query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, table.LogicalName);
            }

            // "createdon" is the logical name of the attribute displayed as "Changed Date" on the
            // audit table. Conditions must use the logical name; Dataverse resolves the localized
            // label itself when the Bulk Delete system job is rendered.
            query.Criteria.AddCondition("createdon", ConditionOperator.GreaterEqual, fromUtc);
            query.Criteria.AddCondition("createdon", ConditionOperator.LessThan, toUtcExclusive);

            return query;
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            var selectedTables = clbTables.CheckedItems.Cast<TableItem>().ToList();
            if (selectedTables.Count == 0)
            {
                MessageBox.Show(this, "Select at least one table.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DateTime fromUtc, toUtc;
            if (!TryGetUtcRange(out fromUtc, out toUtc))
            {
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                btnPreview.Enabled = false;

                var sb = new System.Text.StringBuilder();
                sb.AppendFormat("Preview - Changed Date from {0:yyyy-MM-dd} 00:00 UTC (incl.) to {1:yyyy-MM-dd} 00:00 UTC (excl.):",
                    fromUtc, toUtc);

                foreach (var table in selectedTables)
                {
                    int count = 0;
                    DateTime? min = null;
                    DateTime? max = null;
                    var typeCodes = new HashSet<string>();

                    var query = BuildAuditQuery(table, fromUtc, toUtc,
                        new ColumnSet("createdon", "objecttypecode"));
                    query.PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 };

                    while (true)
                    {
                        var page = service.RetrieveMultiple(query);
                        foreach (var record in page.Entities)
                        {
                            count++;

                            var createdOn = record.GetAttributeValue<DateTime>("createdon");
                            if (createdOn != DateTime.MinValue)
                            {
                                if (min == null || createdOn < min.Value) min = createdOn;
                                if (max == null || createdOn > max.Value) max = createdOn;
                            }

                            var otc = record.Contains("objecttypecode")
                                ? record["objecttypecode"].ToString()
                                : "(null)";
                            typeCodes.Add(otc);
                        }

                        if (!page.MoreRecords || count >= 200000)
                        {
                            break;
                        }

                        query.PageInfo.PageNumber++;
                        query.PageInfo.PagingCookie = page.PagingCookie;
                    }

                    var countText = count >= 200000 ? "200000+" : count.ToString();
                    sb.AppendLine();
                    sb.AppendFormat("- {0} (otc {1}): {2} record(s)",
                        table.LogicalName,
                        table.ObjectTypeCode.HasValue ? table.ObjectTypeCode.Value.ToString() : "n/a",
                        countText);
                    if (count > 0)
                    {
                        sb.AppendFormat(" | Changed Date {0:yyyy-MM-dd HH:mm} .. {1:yyyy-MM-dd HH:mm} UTC | objecttypecode: {2}",
                            min.Value.ToUniversalTime(), max.Value.ToUniversalTime(),
                            string.Join(", ", typeCodes));
                    }
                    else
                    {
                        sb.Append(AppendNoMatchDiagnostic(table));
                    }
                }

                lblResult.ForeColor = System.Drawing.Color.Black;
                lblResult.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                lblResult.ForeColor = System.Drawing.Color.DarkRed;
                lblResult.Text = "Preview failed: " + ex.Message;
                MessageBox.Show(this, "An error occured: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                btnPreview.Enabled = true;
            }
        }

        /// <summary>
        /// When the selected range yields no record, reports whether the table has any audit row
        /// at all and the oldest/newest available createdon, so an empty result can be told apart
        /// from a wrong date range or a wrong entity filter.
        /// </summary>
        private string AppendNoMatchDiagnostic(TableItem table)
        {
            try
            {
                var probe = new QueryExpression("audit")
                {
                    ColumnSet = new ColumnSet("createdon"),
                    Criteria = new FilterExpression(LogicalOperator.And),
                    TopCount = 1
                };

                if (table.ObjectTypeCode.HasValue)
                {
                    probe.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, table.ObjectTypeCode.Value);
                }
                else
                {
                    probe.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, table.LogicalName);
                }

                probe.AddOrder("createdon", OrderType.Ascending);
                var oldest = service.RetrieveMultiple(probe).Entities.FirstOrDefault();

                if (oldest == null)
                {
                    return " | no audit record exists for this table (check that auditing is enabled and data has changed)";
                }

                probe.Orders.Clear();
                probe.AddOrder("createdon", OrderType.Descending);
                var newest = service.RetrieveMultiple(probe).Entities.FirstOrDefault();

                return string.Format(" | table has audit data with Changed Date from {0:yyyy-MM-dd HH:mm} to {1:yyyy-MM-dd HH:mm} UTC - adjust the range",
                    oldest.GetAttributeValue<DateTime>("createdon").ToUniversalTime(),
                    newest.GetAttributeValue<DateTime>("createdon").ToUniversalTime());
            }
            catch (Exception ex)
            {
                return " | diagnostic failed: " + ex.Message;
            }
        }

        private class TableItem
        {
            public TableItem(string logicalName, string displayName, int? objectTypeCode)
            {
                LogicalName = logicalName;
                DisplayName = displayName;
                ObjectTypeCode = objectTypeCode;
            }

            public string LogicalName { get; }

            public string DisplayName { get; }

            public int? ObjectTypeCode { get; }

            public override string ToString()
            {
                return string.Format("{0} ({1})", DisplayName, LogicalName);
            }
        }
    }
}
