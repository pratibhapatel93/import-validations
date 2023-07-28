using System.ComponentModel.DataAnnotations;

namespace ImportData.Models
{
    public class Countries
    {
        [Key]
        public string CountryID { get; set; }
        public string CountryName { get; set; }
        public string TwoCharCountryCode { get; set; }
        public string ThreeCharCountryCode { get; set; }

    }
}
