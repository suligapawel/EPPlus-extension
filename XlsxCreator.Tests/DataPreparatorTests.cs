using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;
using XlsxCreator.Tests.Fixtures;

namespace XlsxCreator.Tests
{
    public class DataPreparatorTests
    {
        private IEnumerable<FakeData> _fakeData;
        private DataPreparator<FakeData> _dataPreparator;

        [SetUp]
        public void Setup()
        {
            _fakeData = Fake.GetFakeData();
            _dataPreparator = new DataPreparator<FakeData>(_fakeData);
        }

        [Test]
        public void Should_return_prepared_collection()
        {
            var result = _dataPreparator.Prepare();

            Assert.That(result.Count, Is.EqualTo(_fakeData.Count()));
        }

        [Test]
        public void Should_return_prepared_collection_with_headers()
        {
            var result = _dataPreparator.PrepareWithHeaders();

            Assert.That(result.Count, Is.EqualTo(_fakeData.Count() + 1));
        }

        [Test]
        public void Should_return_sorted_collection_with_id_placed_in_first_place_when_id_has_attribute_with_the_biggest_priority()
        {
            var result = _dataPreparator.PrepareWithHeaders();

            Assert.That(result.First().First(), Is.EqualTo("Employee Id"));
            Assert.That(result.Skip(1).First().First(), Is.EqualTo(_fakeData.First().Id));
        }

        [Test]
        public void Should_return_collection_without_property_which_has_not_attribute()
        {
            var result = _dataPreparator.Prepare();

            Assert.That(!result.Any(x => x.Any(y => y == _fakeData.First().PhoneNumber)));
        }
    }
}