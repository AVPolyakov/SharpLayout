using System.Linq;
using DataModels;

namespace Examples
{
    public class PaymentOrderQuery
    {
        public PaymentOrderData Get(int id)
        {
            using (var db = new SharpLayoutDB())
            {
                return db.PaymentOrders
                    .Where(o => o.Id == id)
                    .Select(o => new PaymentOrderData {
                        IncomingDate = o.IncomingDate,
                        OutcomingDate = o.OutcomingDate,
                    })
                    .Single();
            }
        }
    }
}