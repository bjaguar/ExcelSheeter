using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelSheeter.Test
{
    [TestClass]
    public class StyleCollectionTest
    {
        [TestMethod]
        public void StyleCollection_CreateNew()
        {
            // Arrange

            // Act
            var list = new StyleCollection();

            // Assert
            Assert.IsNotNull(list);
        }

        [TestMethod]
        public void StyleCollection_Add()
        {
            // Arrange
            var style = new Style();
            var list = new StyleCollection();

            // Act
            list.Add(style);

            // Assert
            Assert.AreEqual(1, list.Count);
        }

        [TestMethod]
        public void StyleCollection_Clear()
        {
            // Arrange
            var style = new Style();
            var list = new StyleCollection();
            list.Add(style);

            // Act
            list.Clear();

            // Assert
            Assert.AreEqual(0, list.Count);
        }

        [TestMethod]
        public void StyleCollection_Contains()
        {
            // Arrange
            var style = new Style();
            var list = new StyleCollection();
            list.Add(style);

            // Act
            var value = list.Contains(style);

            // Assert
            Assert.IsTrue(value);
        }

        [TestMethod]
        public void StyleCollection_Remove()
        {
            // Arrange
            var style = new Style();
            var list = new StyleCollection();
            list.Add(style);

            // Act
            var value = list.Remove(style);

            // Assert
            Assert.AreEqual(0, list.Count);
            Assert.IsTrue(value);
        }

        [TestMethod]
        public void StyleCollection_OuterXml()
        {
            // Arrange
            var style = new Style();
            var list = new StyleCollection();
            list.Add(style);

            // Act
            var xml = list.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }
    }
}
