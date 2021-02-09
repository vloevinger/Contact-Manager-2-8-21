using System.Web;
using System.Web.Mvc;

namespace ContactManagerCsharp2_7_2021
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
