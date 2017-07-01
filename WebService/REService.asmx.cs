using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.Services.Description;
using System.Web.Services.Protocols;
using Excel = Microsoft.Office.Interop.Excel;

namespace WebService
{
    /// <summary>
    /// REService 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
    // [System.Web.Script.Services.ScriptService]
    public class REService : System.Web.Services.WebService
    {
        /// <summary>
        /// 卡方计算
        /// </summary>
        /// <param name="val1">参数1</param>
        /// <param name="val2">参数2</param>
        /// <returns>计算结果</returns>
        [WebMethod]
        [SoapRpcMethod(Use = SoapBindingUse.Literal, Action = "http://tempuri.org/ChiInv", RequestNamespace = "http://tempuri.org/", ResponseNamespace = "http://tempuri.org/")]
        public double ChiInv(double val1, double val2)
        {
            Excel.Application app = null;

            try
            {
                app = new Excel.Application();
                var val = app.WorksheetFunction.ChiInv(val1, val2);
                return val;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                app.Quit();
            }
        }
    }
}
