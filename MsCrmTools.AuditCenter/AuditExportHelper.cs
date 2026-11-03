using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace MsCrmTools.AuditCenter
{
    internal static class AuditExportHelper
    {
        private static readonly string[] EntityHeaders =
        {
            "Display Name",
            "Logical Name",
            "Schema Name",
            "Object Type Code",
            "Audit Enabled",
            "Primary ID Attribute",
            "Primary Name Attribute",
            "Ownership Type",
            "Custom Entity",
            "Description"
        };

        private static readonly string[] AttributeHeaders =
        {
            "Entity Display Name",
            "Entity Logical Name",
            "Entity Schema Name",
            "Display Name",
            "Logical Name",
            "Schema Name",
            "Attribute Type",
            "Audit Enabled",
            "Primary ID",
            "Required Level",
            "Valid for Create",
            "Valid for Update",
            "Custom Attribute",
            "Description"
        };

        private static readonly string[] CsvHeaders =
        {
            "Record Type",
            "Entity Display Name",
            "Entity Logical Name",
            "Entity Schema Name",
            "Object Type Code",
            "Entity Audit Enabled",
            "Primary ID Attribute",
            "Primary Name Attribute",
            "Ownership Type",
            "Custom Entity",
            "Entity Description",
            "Attribute Display Name",
            "Attribute Logical Name",
            "Attribute Schema Name",
            "Attribute Type",
            "Attribute Audit Enabled",
            "Primary ID",
            "Required Level",
            "Valid for Create",
            "Valid for Update",
            "Custom Attribute",
            "Attribute Description"
        };

        public static void ExportCsv(IEnumerable<EntityMetadata> metadata, string path)
        {
            var rows = CreateCsvRows(metadata);
            using (var writer = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                writer.WriteLine(string.Join(",", CsvHeaders.Select(EscapeCsv)));
                foreach (var row in rows)
                {
                    writer.WriteLine(string.Join(",", row.Select(EscapeCsv)));
                }
            }
        }

        public static void ExportExcel(IEnumerable<EntityMetadata> metadata, string path)
        {
            var entities = GetAuditedEntities(metadata);
            var entityRows = entities.Select(entity => new[]
            {
                GetLabel(entity.DisplayName),
                entity.LogicalName,
                entity.SchemaName,
                Format(entity.ObjectTypeCode),
                Format(entity.IsAuditEnabled),
                entity.PrimaryIdAttribute,
                entity.PrimaryNameAttribute,
                Format(entity.OwnershipType),
                Format(entity.IsCustomEntity),
                GetLabel(entity.Description)
            }).ToList();

            var attributeRows = entities
                .SelectMany(entity => GetAuditedAttributes(entity).Select(attribute => new[]
                {
                    GetLabel(entity.DisplayName),
                    entity.LogicalName,
                    entity.SchemaName,
                    GetLabel(attribute.DisplayName),
                    attribute.LogicalName,
                    attribute.SchemaName,
                    GetAttributeType(attribute),
                    Format(attribute.IsAuditEnabled),
                    Format(attribute.IsPrimaryId),
                    attribute.RequiredLevel == null ? string.Empty : Format(attribute.RequiredLevel.Value),
                    Format(attribute.IsValidForCreate),
                    Format(attribute.IsValidForUpdate),
                    Format(attribute.IsCustomAttribute),
                    GetLabel(attribute.Description)
                }))
                .ToList();

            using (var stream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteEntry(archive, "[Content_Types].xml", writer => WriteContentTypes(writer));
                WriteEntry(archive, "_rels/.rels", writer => WriteRootRelationships(writer));
                WriteEntry(archive, "xl/workbook.xml", writer => WriteWorkbook(writer));
                WriteEntry(archive, "xl/_rels/workbook.xml.rels", writer => WriteWorkbookRelationships(writer));
                WriteEntry(archive, "xl/worksheets/sheet1.xml",
                    writer => WriteWorksheet(writer, EntityHeaders, entityRows));
                WriteEntry(archive, "xl/worksheets/sheet2.xml",
                    writer => WriteWorksheet(writer, AttributeHeaders, attributeRows));
            }
        }

        private static IEnumerable<string[]> CreateCsvRows(IEnumerable<EntityMetadata> metadata)
        {
            foreach (var entity in GetAuditedEntities(metadata))
            {
                var attributes = GetAuditedAttributes(entity).ToList();
                if (attributes.Count == 0)
                {
                    yield return CreateCsvRow(entity, null);
                    continue;
                }

                foreach (var attribute in attributes)
                {
                    yield return CreateCsvRow(entity, attribute);
                }
            }
        }

        private static string[] CreateCsvRow(EntityMetadata entity, AttributeMetadata attribute)
        {
            return new[]
            {
                attribute == null ? "Table" : "Column",
                GetLabel(entity.DisplayName),
                entity.LogicalName,
                entity.SchemaName,
                Format(entity.ObjectTypeCode),
                Format(entity.IsAuditEnabled),
                entity.PrimaryIdAttribute,
                entity.PrimaryNameAttribute,
                Format(entity.OwnershipType),
                Format(entity.IsCustomEntity),
                GetLabel(entity.Description),
                attribute == null ? null : GetLabel(attribute.DisplayName),
                attribute == null ? null : attribute.LogicalName,
                attribute == null ? null : attribute.SchemaName,
                attribute == null ? null : GetAttributeType(attribute),
                attribute == null ? null : Format(attribute.IsAuditEnabled),
                attribute == null ? null : Format(attribute.IsPrimaryId),
                attribute == null || attribute.RequiredLevel == null
                    ? null
                    : Format(attribute.RequiredLevel.Value),
                attribute == null ? null : Format(attribute.IsValidForCreate),
                attribute == null ? null : Format(attribute.IsValidForUpdate),
                attribute == null ? null : Format(attribute.IsCustomAttribute),
                attribute == null ? null : GetLabel(attribute.Description)
            };
        }

        private static List<EntityMetadata> GetAuditedEntities(IEnumerable<EntityMetadata> metadata)
        {
            if (metadata == null)
            {
                throw new ArgumentNullException("metadata");
            }

            return metadata
                .Where(entity => entity != null && entity.IsAuditEnabled != null && entity.IsAuditEnabled.Value)
                .OrderBy(entity => entity.LogicalName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(entity => entity.LogicalName, StringComparer.Ordinal)
                .ToList();
        }

        private static IEnumerable<AttributeMetadata> GetAuditedAttributes(EntityMetadata entity)
        {
            return (entity.Attributes ?? new AttributeMetadata[0])
                .Where(attribute => attribute != null
                    && attribute.IsAuditEnabled != null
                    && attribute.IsAuditEnabled.Value
                    && attribute.AttributeOf == null)
                .OrderBy(attribute => attribute.LogicalName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(attribute => attribute.LogicalName, StringComparer.Ordinal);
        }

        private static string GetAttributeType(AttributeMetadata attribute)
        {
            return attribute.AttributeTypeName != null && !string.IsNullOrEmpty(attribute.AttributeTypeName.Value)
                ? attribute.AttributeTypeName.Value
                : Format(attribute.AttributeType);
        }

        private static string GetLabel(Label label)
        {
            return label == null || label.UserLocalizedLabel == null
                ? string.Empty
                : label.UserLocalizedLabel.Label ?? string.Empty;
        }

        private static string Format(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            var managedProperty = value as Microsoft.Xrm.Sdk.BooleanManagedProperty;
            if (managedProperty != null)
            {
                return managedProperty.Value.ToString(CultureInfo.InvariantCulture);
            }

            var enumValue = value as Enum;
            if (enumValue != null)
            {
                return enumValue.ToString();
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string EscapeCsv(string value)
        {
            value = value ?? string.Empty;
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0)
            {
                return value;
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static void WriteEntry(ZipArchive archive, string entryName, Action<XmlWriter> writeXml)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = true
            }))
            {
                writeXml(writer);
            }
        }

        private static void WriteContentTypes(XmlWriter writer)
        {
            const string ns = "http://schemas.openxmlformats.org/package/2006/content-types";
            writer.WriteStartDocument();
            writer.WriteStartElement("Types", ns);
            WriteDefault(writer, ns, "rels", "application/vnd.openxmlformats-package.relationships+xml");
            WriteDefault(writer, ns, "xml", "application/xml");
            WriteOverride(writer, ns, "/xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            WriteOverride(writer, ns, "/xl/worksheets/sheet1.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            WriteOverride(writer, ns, "/xl/worksheets/sheet2.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private static void WriteDefault(XmlWriter writer, string ns, string extension, string contentType)
        {
            writer.WriteStartElement("Default", ns);
            writer.WriteAttributeString("Extension", extension);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        private static void WriteOverride(XmlWriter writer, string ns, string partName, string contentType)
        {
            writer.WriteStartElement("Override", ns);
            writer.WriteAttributeString("PartName", partName);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        private static void WriteRootRelationships(XmlWriter writer)
        {
            const string ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            writer.WriteStartDocument();
            writer.WriteStartElement("Relationships", ns);
            writer.WriteStartElement("Relationship", ns);
            writer.WriteAttributeString("Id", "rId1");
            writer.WriteAttributeString("Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            writer.WriteAttributeString("Target", "xl/workbook.xml");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private static void WriteWorkbook(XmlWriter writer)
        {
            const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string relationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            writer.WriteStartDocument();
            writer.WriteStartElement("workbook", ns);
            writer.WriteAttributeString("xmlns", "r", null, relationshipsNs);
            writer.WriteStartElement("sheets", ns);
            WriteSheet(writer, "Audited Tables", 1, "rId1");
            WriteSheet(writer, "Audited Columns", 2, "rId2");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private static void WriteSheet(XmlWriter writer, string name, int sheetId, string relationshipId)
        {
            const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            writer.WriteStartElement("sheet", ns);
            writer.WriteAttributeString("name", name);
            writer.WriteAttributeString("sheetId", sheetId.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("r", "id",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships", relationshipId);
            writer.WriteEndElement();
        }

        private static void WriteWorkbookRelationships(XmlWriter writer)
        {
            const string ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            writer.WriteStartDocument();
            writer.WriteStartElement("Relationships", ns);
            WriteWorksheetRelationship(writer, "rId1", "worksheets/sheet1.xml");
            WriteWorksheetRelationship(writer, "rId2", "worksheets/sheet2.xml");
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private static void WriteWorksheetRelationship(XmlWriter writer, string id, string target)
        {
            const string ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            writer.WriteStartElement("Relationship", ns);
            writer.WriteAttributeString("Id", id);
            writer.WriteAttributeString("Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            writer.WriteAttributeString("Target", target);
            writer.WriteEndElement();
        }

        private static void WriteWorksheet(XmlWriter writer, string[] headers, IEnumerable<string[]> rows)
        {
            const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            writer.WriteStartDocument();
            writer.WriteStartElement("worksheet", ns);
            writer.WriteStartElement("sheetData", ns);
            WriteRow(writer, ns, 1, headers);
            var rowNumber = 2;
            foreach (var row in rows)
            {
                WriteRow(writer, ns, rowNumber++, row);
            }
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private static void WriteRow(XmlWriter writer, string ns, int rowNumber, string[] values)
        {
            writer.WriteStartElement("row", ns);
            writer.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
            for (var index = 0; index < values.Length; index++)
            {
                writer.WriteStartElement("c", ns);
                writer.WriteAttributeString("r", GetCellReference(index + 1, rowNumber));
                writer.WriteAttributeString("t", "inlineStr");
                writer.WriteStartElement("is", ns);
                writer.WriteStartElement("t", ns);
                writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                writer.WriteString(values[index] ?? string.Empty);
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private static string GetCellReference(int column, int row)
        {
            var columnName = string.Empty;
            while (column > 0)
            {
                column--;
                columnName = (char)('A' + column % 26) + columnName;
                column /= 26;
            }

            return columnName + row.ToString(CultureInfo.InvariantCulture);
        }
    }
}
