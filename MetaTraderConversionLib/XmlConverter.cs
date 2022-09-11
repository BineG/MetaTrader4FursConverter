using MetaTraderConversionLib.Dtos;
using MetaTraderConversionLib.Xml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace MetaTraderConversionLib
{
    public class XmlConverter
    {
        private readonly ConversionInfoDto _info;
        private readonly ArchiveParser _parser;

        public XmlConverter(ConversionInfoDto info)
        {
            _info = info;
            _parser = new ArchiveParser();
        }

        public void Convert(string sourceFile, string targetFile)
        {
            XmlDocument doc = _parser.Parse(sourceFile);
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            // search inside namespace Workbook
            nsmgr.AddNamespace("wb", "urn:schemas-microsoft-com:office:spreadsheet");

            XmlNode? table = doc.DocumentElement?.SelectSingleNode("//wb:Table", nsmgr);

            if (table == null)
                throw new ArgumentNullException(nameof(table));

            List<XmlNode> rows = table.ChildNodes
                .Cast<XmlNode>()
                .Where(x => x.Name == "Row")
                .ToList();

            List<TransactionDto> transactions = new List<TransactionDto>();
            bool transactionsStarted = false;
            foreach (XmlNode row in rows)
            {
                List<XmlNode> cells = row.ChildNodes.Cast<XmlNode>().ToList();
                if (cells.FirstOrDefault()?.InnerText == "Time")
                {
                    transactionsStarted = true;
                    continue;
                }

                if (!transactionsStarted)
                    continue;

                if (cells.ElementAt(0).InnerText == "Orders")
                {
                    // end of data
                    break;
                }

                transactions.Add(new TransactionDto()
                {
                    BuyTime = ParseTime(cells.ElementAt(0)),
                    AssetName = cells.ElementAt(2).InnerText,
                    Type = cells.ElementAt(3).InnerText,
                    Amount = ParseNumber(cells.ElementAt(4), 4),
                    BuyPrice = ParseNumber(cells.ElementAt(5), 4),
                    SellTime = ParseTime(cells.ElementAt(8)),
                    SellPrice = ParseNumber(cells.ElementAt(9), 4),
                    Profit = ParseNumber(cells.ElementAt(12), 4)
                });
            }

            var assetGroups = transactions.GroupBy(x => x.AssetName).ToList();

            Envelope rootDoc = new Envelope()
            {
                Header = new Header
                {
                    taxpayer = new taxPayerType()
                    {
                        taxpayerType = taxpayerTypeType.FO,
                        taxpayerTypeSpecified = true,
                        address1 = _info.PersonInfo.Address1,
                        city = _info.PersonInfo.City,
                        postNumber = _info.PersonInfo.PostNumber,
                        birthDate = _info.PersonInfo.Birthdate,
                        birthDateSpecified = true,
                        name = _info.PersonInfo.NameSurname,
                        ItemElementName = ItemChoiceType.taxNumber,
                        Item = _info.PersonInfo.TaxNumber
                    }
                },
                body = new EnvelopeBody()
                {
                    bodyContent = new object(),
                    Doh_KDVP = new Doh_KDVP()
                    {
                        KDVP = new Doh_KDVPKDVP()
                        {
                            DocumentWorkflowID = _info.DocumentInfo.DocumentType,
                            Year = 2020,
                            YearSpecified = true,
                            PeriodStart = new DateTime(2020, 1, 1),
                            PeriodStartSpecified = true,
                            PeriodEnd = new DateTime(2020, 12, 31),
                            PeriodEndSpecified = true,
                            IsResident = true,
                            IsResidentSpecified = true,
                            TelephoneNumber = _info.PersonInfo.Telephone,
                            Email = _info.PersonInfo.Email,
                            // stevilo razlicnih trgovalnih papirjev
                            SecurityCount = assetGroups.Count
                        },
                        KDVPItem = assetGroups.Select((g, index) =>
                        {
                            List<SecuritiesRow> securitiesRows = new();
                            int id = 0;
                            for (int entryIndex = 0; entryIndex < g.Count(); entryIndex++)
                            {
                                var entry = g.ElementAt(entryIndex);

                                securitiesRows.Add(new SecuritiesRow()
                                {
                                    ID = id++,
                                    Item = new SecuritiesRowPurchase()
                                    {
                                        F1 = entry.BuyTime,
                                        F1Specified = true,
                                        F2 = typeGainType.A,
                                        F2Specified = true,
                                        F3 = entry.Amount,
                                        F3Specified = true,
                                        F4 = entry.BuyPrice,
                                        F4Specified = true,
                                    },
                                    F8 = entry.Amount,
                                    F8Specified = true
                                });
                                securitiesRows.Add(new SecuritiesRow()
                                {
                                    ID = id++,
                                    Item = new SecuritiesRowSale()
                                    {
                                        F6 = entry.SellTime,
                                        F6Specified = true,
                                        F7 = entry.Amount,
                                        F7Specified = true,
                                        F9 = entry.SellPrice,
                                        F9Specified = true,
                                        F10 = true,
                                        F10Specified = true
                                    },
                                    F8 = 0,
                                    F8Specified = true
                                });
                            }

                            return new Doh_KDVPKDVPItem()
                            {
                                ItemID = index + 1,
                                ItemIDSpecified = true,
                                InventoryListType = typeInventory.PLVP,
                                HasForeignTax = false,
                                HasForeignTaxSpecified = true,
                                HasLossTransfer = false,
                                HasLossTransferSpecified = true,
                                ForeignTransfer = false,
                                ForeignTransferSpecified = true,
                                TaxDecreaseConformance = false,
                                TaxDecreaseConformanceSpecified = true,
                                Name = g.Key,
                                Item = new Securities()
                                {
                                    Name = g.Key,
                                    IsFond = false,
                                    Row = securitiesRows.ToArray()
                                }
                            };
                        })
                        .ToArray()
                    }
                },
                Signatures = new Signatures(),
                AttachmentList = new AttachmentListExternalAttachment[0]
            };

            XmlSerializer serializer = new XmlSerializer(typeof(Envelope));
            TextWriter writer = new StreamWriter(targetFile);
            serializer.Serialize(writer, rootDoc);
            writer.Close();
        }

        private DateTime ParseTime(XmlNode cell)
        {
            return DateTime.ParseExact(cell.InnerText, "yyyy.MM.dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private decimal ParseNumber(XmlNode cell, int decimals)
        {
            decimal d = decimal.Parse(cell.InnerText, CultureInfo.InvariantCulture);
            return decimal.Parse(d.ToString($"n{decimals}", CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture);
        }
    }
}
