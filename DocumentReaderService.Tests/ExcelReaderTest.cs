﻿using System;
using System.Collections.Generic;
using System.IO;
using DocumentReaderService.Exceptions;
using NUnit.Framework;
using FluentAssertions;

namespace DocumentReaderService.Tests
{
    [TestFixture]
    public class ExcelReaderTest
    {
        private readonly string filePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "TestHelpers\\test.xlsx");
        private readonly Dictionary<string, int> dictionary = new Dictionary<string, int>
        {
            {"Закрытие декабря 2018 года", 1},
            {"Входящий акт приемки услуг", 8},
            {"Бухгалтерская справка", 5},
            {"Исходящий акт приемки услуг", 2},
            {"Отчет по безналичной рознице", 2},
            {"Банковский ордер", 39},
            {"Входящее платежное поручение", 53},
            {"Исходящее платежное поручение", 31},
            {"Уплата налогов и взносов", 2},
            {"Входящий счет на оплату", 8},
            {"Закрытие ноября 2018 года", 1}
        };

        [Test]
        public void CanReadFromFile()
        {
            var res = ExcelReader.ReadFromFile(filePath);
            res.Should().BeOfType(typeof(Dictionary<string, int>));
        }

        [Test]
        public void ReadFromFileCheckResult()
        {
            var resultDictionary = ExcelReader.ReadFromFile(filePath);
            resultDictionary.Should().BeEquivalentTo(dictionary);
        }

        [Test]
        public void CanReadByteArray()
        {
            var bytes = File.ReadAllBytes(filePath);
            var res = ExcelReader.ReadFromByteArray(bytes);
            res.Should().BeOfType(typeof(Dictionary<string, int>));
        }

        [Test]
        public void ReadFromByteArrayCheckResult()
        {
            var bytes = File.ReadAllBytes(filePath);
            var resultDictionary = ExcelReader.ReadFromByteArray(bytes);
            resultDictionary.Should().BeEquivalentTo(dictionary);
        }

        [Test]
        public void CanThrowExceptionExcelReader()
        {
            var invalidFilePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "TestHelpers\\invalid_format.xlsx");
            Action act = () => ExcelReader.ReadFromFile(invalidFilePath);
            act.Should().Throw<ExceptionExcelReader>()
                .WithMessage("Invalid file format");
        }
    }
}
