//------------------------------------------------------------------------------
// <auto-generated>
//    這個程式碼是由範本產生。
//
//    對這個檔案進行手動變更可能導致您的應用程式產生未預期的行為。
//    如果重新產生程式碼，將會覆寫對這個檔案的手動變更。
// </auto-generated>
//------------------------------------------------------------------------------

namespace NorthWind.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    public partial class Customers
    {
        public Customers()
        {
            this.Orders = new HashSet<Orders>();
            this.CustomerDemographics = new HashSet<CustomerDemographics>();
        }

        [DisplayName("客戶編號")]
        public string CustomerID { get; set; }

        [DisplayName("公司名稱")]
        public string CompanyName { get; set; }

        [DisplayName("客戶名稱")]
        public string ContactName { get; set; }

        [DisplayName("客戶職稱")]
        public string ContactTitle { get; set; }

        [DisplayName("地址")]
        public string Address { get; set; }

        [DisplayName("城市")]
        public string City { get; set; }

        [DisplayName("地區")]
        public string Region { get; set; }

        [DisplayName("郵遞區號")]
        public string PostalCode { get; set; }

        [DisplayName("國家")]
        public string Country { get; set; }

        [DisplayName("電話")]
        public string Phone { get; set; }

        [DisplayName("傳真")]
        public string Fax { get; set; }
    
        public virtual ICollection<Orders> Orders { get; set; }
        public virtual ICollection<CustomerDemographics> CustomerDemographics { get; set; }
    }
}