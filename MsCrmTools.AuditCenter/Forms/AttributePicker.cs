using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml;

namespace MsCrmTools.AuditCenter.Forms
{
    public partial class AttributePicker : Form
    {
        public List<AttributeMetadata> AttributesToAdd;
        private readonly IEnumerable<string> alreadySelectedAttributes;
        private readonly List<EntityMetadata> emds;
        private readonly IOrganizationService service;

        /// <summary>
        /// Initializes a new instance of <see cref="AttributePicker"/> for one or more tables
        /// </summary>
        /// <param name="emds">Tables whose attributes can be enabled for audit</param>
        /// <param name="alreadySelectedAttributes">Keys, formatted as 'tableLogicalName.attributeLogicalName', of attributes already audited</param>
        /// <param name="service">Organization service</param>
        public AttributePicker(IEnumerable<EntityMetadata> emds, IEnumerable<string> alreadySelectedAttributes, IOrganizationService service)
        {
            this.emds = emds.ToList();
            this.alreadySelectedAttributes = alreadySelectedAttributes;
            this.service = service;
            InitializeComponent();
        }

        private void AttributePickerLoad(object sender, EventArgs e)
        {
            var useGroups = emds.Count > 1;
            lvAttributes.ShowGroups = useGroups;

            foreach (var emd in emds)
            {
                XmlDocument allFormsDoc = MetadataHelper.RetrieveEntityForms(emd.LogicalName, service);

                ListViewGroup group = null;
                if (useGroups)
                {
                    group = new ListViewGroup(emd.LogicalName,
                        string.Format("{0} ({1})", emd.DisplayName?.UserLocalizedLabel?.Label ?? "N/A", emd.LogicalName));
                    lvAttributes.Groups.Add(group);
                }

                foreach (AttributeMetadata amd in emd.Attributes.Where(a =>
                    !alreadySelectedAttributes.Contains(string.Concat(emd.LogicalName, ".", a.LogicalName))
                    && a.AttributeOf == null))
                {
                    string displayName = amd.DisplayName?.UserLocalizedLabel?.Label ?? "N/A";

                    var item = new ListViewItem { Text = displayName, Tag = amd };
                    item.SubItems.Add(amd.LogicalName);
                    item.SubItems.Add(emd.LogicalName);
                    item.SubItems.Add((allFormsDoc.SelectSingleNode("//control[@datafieldname='" + amd.LogicalName + "']") != null).ToString());

                    if (group != null)
                    {
                        item.Group = group;
                    }

                    lvAttributes.Items.Add(item);
                }
            }
        }

        private void BtnCancelClick(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void BtnCheckAttrOnFormsClick(object sender, EventArgs e)
        {
            foreach (ListViewItem item in lvAttributes.Items)
            {
                item.Checked = item.SubItems[3].Text.ToLower() == "true";
            }
        }

        private void BtnCheckClick(object sender, EventArgs e)
        {
            foreach (ListViewItem item in lvAttributes.Items)
            {
                item.Checked = ((Button)sender).Text == "Check All";
            }

            ((Button)sender).Text = ((Button)sender).Text == "Check All" ? "Clear All" : "Check All";
        }

        private void BtnOkClick(object sender, EventArgs e)
        {
            AttributesToAdd = new List<AttributeMetadata>();

            foreach (ListViewItem item in lvAttributes.CheckedItems)
            {
                AttributesToAdd.Add((AttributeMetadata)item.Tag);
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void ListViewColumnClick(object sender, ColumnClickEventArgs e)
        {
            var lv = (ListView)sender;

            lv.Sorting = lv.Sorting == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;

            lv.ListViewItemSorter = new ListViewItemComparer(e.Column, lv.Sorting);
        }
    }
}