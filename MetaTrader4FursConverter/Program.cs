// See https://aka.ms/new-console-template for more information
using MetaTraderConversionLib;
using MetaTraderConversionLib.Dtos;
using Microsoft.Extensions.Configuration;

IConfiguration config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .AddJsonFile("appsettings.secret.json", true)
    .AddEnvironmentVariables()
    .Build();

ConversionInfoDto conversionInfo = config.GetSection("ConversionInfo").Get<ConversionInfoDto>();

XmlConverter converter = new XmlConverter(conversionInfo);

converter.Convert("ReportHistory-543742.xml", "edavki");