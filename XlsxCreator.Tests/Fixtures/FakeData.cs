using System;
using System.Collections.Generic;
using XlsxCreator.Attributes;

namespace XlsxCreator.Tests.Fixtures
{
    internal class FakeData
    {
        [XlsxTable("Employee name", 3)]
        public string Name { get; set; }

        [XlsxTable("Employee city")]
        public string City { get; set; }

        [XlsxTable("Lastname", 5)]
        public string LastName { get; set; }

        [XlsxTable("Address", 2)]
        public string Address { get; set; }

        public string PhoneNumber { get; set; }

        [XlsxTable("Employee Id", 1)]
        public int Id { get; set; }

        [XlsxTable("Birthday", 55)]
        public DateTime Birthday { get; set; }
    }

    internal static class Fake
    {
        public static IEnumerable<FakeData> GetFakeData()
            => new FakeData[]
            {
                new FakeData
                {
                    Id = 33,
                    Name = "Paul",
                    LastName = "Kowalsky",
                    City = "New York",
                    PhoneNumber = "999999999",
                    Birthday = DateTime.Now
                },
                new FakeData
                {
                    Id = 21,
                    Name = "John",
                    LastName = "Johnson",
                    City = "Chicago",
                    PhoneNumber = "888888888"
                }
            };
    }
}
