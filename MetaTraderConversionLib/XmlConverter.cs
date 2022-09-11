using MetaTraderConversionLib.Dtos;
using MetaTraderConversionLib.Xml;
using MetaTraderConversionLib.Xml.DIFI;
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

            //transactions = transactions.Where(t => t.AssetName == "NatGas" /*&& t.Type == "sell"*/).ToList();

            //var natgas = transactions.Where(t => t.AssetName == "NatGas").Sum(x => x.Profit);

            //var assetGroups = transactions.GroupBy(x => new { x.AssetName, SellYear = x.SellTime.Year }).ToList();

            OutputD_IFI(targetFile, transactions);
            OutputDoh_KDVP(targetFile, transactions);
        }

        private void OutputDoh_KDVP(string targetFile, List<TransactionDto> transactions)
        {
            List<int> years = transactions
                            .Select(x => x.SellTime.Year)
                            .Distinct()
                            .OrderBy(x => x)
                            .ToList();

            XmlSerializer serializer = new XmlSerializer(typeof(Xml.DohKDVP.Envelope));

            foreach (int currentYear in years)
            {
                var assetGroups = transactions
                    .Where(x => x.SellTime.Year == currentYear)
                    .GroupBy(x => new { x.AssetName, x.Type })
                    .ToList();
                //decimal profit = 0;
                Xml.DohKDVP.Envelope rootDoc = new ()
                {
                    Header = new Xml.DohKDVP.Header
                    {
                        taxpayer = new Xml.DohKDVP.taxPayerType()
                        {
                            taxpayerType = Xml.DohKDVP.taxpayerTypeType.FO,
                            taxpayerTypeSpecified = true,
                            address1 = _info.PersonInfo.Address1,
                            city = _info.PersonInfo.City,
                            postNumber = _info.PersonInfo.PostNumber,
                            birthDate = _info.PersonInfo.Birthdate,
                            birthDateSpecified = true,
                            name = _info.PersonInfo.NameSurname,
                            ItemElementName = Xml.DohKDVP.ItemChoiceType.taxNumber,
                            Item = _info.PersonInfo.TaxNumber,
                        }
                    },
                    AttachmentList = new Xml.DohKDVP.AttachmentListExternalAttachment[0],
                    Signatures = new (),
                    body = new Xml.DohKDVP.EnvelopeBody()
                    {
                        bodyContent = new object(),
                        Doh_KDVP = new ()
                        {
                            KDVP = new Xml.DohKDVP.Doh_KDVPKDVP()
                            {
                                DocumentWorkflowID = _info.DocumentInfo.DocumentType,
                                Year = currentYear,
                                YearSpecified = true,
                                PeriodStart = new DateTime(currentYear, 1, 1),
                                PeriodStartSpecified = true,
                                PeriodEnd = new DateTime(currentYear, 12, 31),
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
                                List<object> securitiesRows = new List<object>();

                                int id = 0;
                                for (int entryIndex = 0; entryIndex < g.Count(); entryIndex++)
                                {
                                    var line = g.ElementAt(entryIndex);

                                    decimal priceDiff = line.BuyPrice - line.SellPrice;

                                    decimal ratio = 1;
                                    if (priceDiff != 0)
                                    {
                                        ratio = Math.Abs(line.Profit / priceDiff);
                                    }
                                    ratio /= line.Amount;

                                    if (g.Key.Type == "buy")
                                    {
                                        //List<Xml.DohKDVP.SecuritiesRow> securitiesRows = new();

                                        securitiesRows.Add(new Xml.DohKDVP.SecuritiesRow()
                                        {
                                            ID = id++,
                                            Item = new Xml.DohKDVP.SecuritiesRowPurchase()
                                            {
                                                F1 = line.BuyTime,
                                                F1Specified = true,
                                                F2 = Xml.DohKDVP.typeGainType.A,
                                                F2Specified = true,
                                                F3 = line.Amount,
                                                F3Specified = true,
                                                F4 = line.BuyPrice,
                                                F4Specified = true,
                                            },
                                            F8 = line.Amount,
                                            F8Specified = true
                                        });
                                        securitiesRows.Add(new Xml.DohKDVP.SecuritiesRow()
                                        {
                                            ID = id++,
                                            Item = new Xml.DohKDVP.SecuritiesRowSale()
                                            {
                                                F6 = line.SellTime,
                                                F6Specified = true,
                                                F7 = line.Amount,
                                                F7Specified = true,
                                                F9 = line.SellPrice,
                                                F9Specified = true,
                                                F10 = true,
                                                F10Specified = true
                                            },
                                            F8 = 0,
                                            F8Specified = true
                                        });
                                    }
                                    else
                                    {
                                        // SELL

                                        securitiesRows.Add(new Xml.DohKDVP.SecuritiesShortRow()
                                        {
                                            ID = id++,
                                            Item = new Xml.DohKDVP.SecuritiesShortRowSale()
                                            {
                                                F6 = line.SellTime,
                                                F6Specified = true,
                                                F7 = line.Amount,
                                                F7Specified = true,
                                                //F9 = line.SellPrice,
                                                F9 = RoundToDecimals(line.BuyPrice * ratio, 4),
                                                F9Specified = true
                                            },
                                            //F8 = 0,
                                            //F8Specified = true
                                        });

                                        securitiesRows.Add(new Xml.DohKDVP.SecuritiesShortRow()
                                        {
                                            ID = id++,
                                            Item = new Xml.DohKDVP.SecuritiesShortRowPurchase()
                                            {
                                                F1 = line.BuyTime,
                                                F1Specified = true,
                                                F2 = Xml.DohKDVP.typeGainType.A,
                                                F2Specified = true,
                                                F3 = line.Amount,
                                                F3Specified = true,
                                                F4 = RoundToDecimals(line.SellPrice * ratio, 4),
                                                F4Specified = true,
                                                
                                            },
                                            //F8 = line.Amount,
                                            //F8Specified = true
                                        });
                                    }
                                }

                                return new Xml.DohKDVP.Doh_KDVPKDVPItem()
                                {
                                    ItemID = index + 1,
                                    ItemIDSpecified = true,
                                    InventoryListType = g.Key.Type == "buy"
                                        ? Xml.DohKDVP.typeInventory.PLVP
                                        : Xml.DohKDVP.typeInventory.PLVPSHORT,
                                    HasForeignTax = false,
                                    HasForeignTaxSpecified = true,
                                    HasLossTransfer = false,
                                    HasLossTransferSpecified = true,
                                    ForeignTransfer = false,
                                    ForeignTransferSpecified = true,
                                    TaxDecreaseConformance = false,
                                    TaxDecreaseConformanceSpecified = true,
                                    Name = g.Key.AssetName,
                                    Item = g.Key.Type == "buy"
                                        ? new Xml.DohKDVP.Securities()
                                        {
                                            Name = g.Key.AssetName,
                                            IsFond = false,
                                            Row = securitiesRows.Select(x => x as Xml.DohKDVP.SecuritiesRow).ToArray()
                                        }
                                        : new Xml.DohKDVP.SecuritiesShort()
                                        {
                                            Name = g.Key.AssetName,
                                            IsFond = false,
                                            Row = securitiesRows.Select(x => x as Xml.DohKDVP.SecuritiesShortRow).ToArray()
                                        }
                                };
                            })
                            .ToArray()
                        },
                    }
                };

                string yearTargetFile = $"KDVP-{targetFile}-{currentYear}.xml";
                using TextWriter writer = new StreamWriter(yearTargetFile);
                serializer.Serialize(writer, rootDoc);

                writer.Close();
            }
        }

        private void OutputD_IFI(string targetFile, List<TransactionDto> transactions)
        {
            List<int> years = transactions
                            .Select(x => x.SellTime.Year)
                            .Distinct()
                            .OrderBy(x => x)
                            .ToList();

            XmlSerializer serializer = new XmlSerializer(typeof(Envelope));

            foreach (int currentYear in years)
            {
                var assetGroups = transactions
                    .Where(x => x.SellTime.Year == currentYear)
                    .GroupBy(x => new { x.AssetName, x.Type })
                    .ToList();
                decimal profit = 0;
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
                        },
                        Workflow = new workflowType
                        {
                            DocumentWorkflowID = _info.DocumentInfo.DocumentType
                        }
                    },
                    body = new EnvelopeBody()
                    {
                        bodyContent = new object(),
                        D_IFI = new D_IFI()
                        {
                            PeriodStart = new DateTime(currentYear, 1, 1),
                            PeriodStartSpecified = true,
                            PeriodEnd = new DateTime(currentYear, 12, 31),
                            PeriodEndSpecified = true,
                            TelephoneNumber = _info.PersonInfo.Telephone,
                            Email = _info.PersonInfo.Email,

                            TItem = assetGroups
                                .Select(g =>
                                {
                                    List<object> items = new List<object>();
                                    foreach (var line in g)
                                    {
                                        decimal priceDiff = line.BuyPrice - line.SellPrice;

                                        decimal ratio = 1;
                                        if (priceDiff != 0)
                                        {
                                            ratio = Math.Abs(line.Profit / priceDiff);
                                        }
                                        ratio /= line.Amount;

                                        if (g.Key.Type == "buy")
                                        {
                                            // buy
                                            var buy = new TSubItem()
                                            {
                                                F8 = line.Amount,
                                                F8Specified = true,
                                                Item = new TSubItemPurchase()
                                                {
                                                    F1 = line.BuyTime,
                                                    F1Specified = true,
                                                    F2 = typeGainType.A,
                                                    F2Specified = true,
                                                    F3 = line.Amount,
                                                    F3Specified = true,
                                                    F4 = RoundToDecimals(line.BuyPrice * ratio, 4),
                                                    //F4 = line.BuyPrice,
                                                    F4Specified = true,
                                                    F9 = ratio > 1,
                                                    F9Specified = true
                                                }
                                            };

                                            items.Add(buy);

                                            // sell
                                            var sell = new TSubItem()
                                            {
                                                F8 = 0,
                                                F8Specified = true,
                                                Item = new TSubItemSale()
                                                {
                                                    F5 = line.SellTime,
                                                    F6 = line.Amount,
                                                    F7 = RoundToDecimals(line.SellPrice * ratio, 4)
                                                    //F7 = line.SellPrice
                                                }
                                            };
                                            items.Add(sell);

                                            decimal calculatedProfit = ((sell.Item as TSubItemSale)?.F7 ?? 0) - ((buy.Item as TSubItemPurchase)?.F4 ?? 0);

                                            profit += calculatedProfit;
                                        }
                                        else
                                        {
                                            // sell
                                            items.Add(new TShortSubItem()
                                            {
                                                F8 = -line.Amount,
                                                F8Specified = true,
                                                Item = new TShortSubItemSale()
                                                {
                                                    F1 = line.BuyTime,
                                                    F1Specified = true,
                                                    F2 = line.Amount,
                                                    F2Specified = true,
                                                    F3 = RoundToDecimals(line.BuyPrice * ratio, 4),
                                                    F3Specified = true,
                                                    F9 = ratio > 1,
                                                    F9Specified = true
                                                }
                                            });

                                            // buy
                                            items.Add(new TShortSubItem()
                                            {
                                                F8 = 0,
                                                F8Specified = true,
                                                Item = new TShortSubItemPurchase()
                                                {
                                                    F4 = line.SellTime,
                                                    F4Specified = true,
                                                    F5 = typeGainType.A,
                                                    F5Specified = true,
                                                    F6 = line.Amount,
                                                    F6Specified = true,
                                                    F7 = RoundToDecimals(line.SellPrice * ratio, 4),
                                                    F7Specified = true
                                                }
                                            });
                                        }
                                    }

                                    return new TItem()
                                    {
                                        TypeId = g.Key.Type == "buy"
                                            ? typeInventory.PLIFI
                                            : typeInventory.PLIFIShort,
                                        Type = "02",
                                        TypeName = "finančne pogodbe na razliko",
                                        Name = g.Key.AssetName,
                                        HasForeignTax = false,
                                        HasForeignTaxSpecified = true,
                                        Items = items.ToArray()
                                    };
                                })
                                .ToArray()
                        },
                    },
                    Signatures = new Signatures(),
                    AttachmentList = new AttachmentListExternalAttachment[0]
                };

                string yearTargetFile = $"DIFI-{targetFile}-{currentYear}.xml";
                using TextWriter writer = new StreamWriter(yearTargetFile);
                serializer.Serialize(writer, rootDoc);

                writer.Close();
            }
        }

        private DateTime ParseTime(XmlNode cell)
        {
            return DateTime.ParseExact(cell.InnerText, "yyyy.MM.dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private decimal ParseNumber(XmlNode cell, int decimals)
        {
            decimal d = decimal.Parse(cell.InnerText, CultureInfo.InvariantCulture);
            return RoundToDecimals(d, decimals);
        }

        private decimal RoundToDecimals(decimal d, int decimals)
        {
            return decimal.Parse(d.ToString($"n{decimals}", CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture);
        }
    }
}
