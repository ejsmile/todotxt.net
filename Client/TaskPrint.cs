using System.Text;
using System.Windows.Controls;
using ToDoLib;

namespace Client
{
    class TaskPrint
    {
        private WebBrowser _browser;
        private mshtml.IHTMLDocument2 _doc;

        public TaskPrint(WebBrowser browser)
        {
            _browser = browser;
            _doc = _browser.Document as mshtml.IHTMLDocument2;
        }

        public void GenerateDocTaskList(ItemCollection tasks)
        {
            _doc.clear();
            _doc.write(Get_PrintContents(tasks));
            _doc.close();
        }

        public void PrintDocTaskList()
        {
            _doc.execCommand("Print", true, 0);
        }

        private string Get_PrintContents(ItemCollection tasks)
        {
            if (tasks == null || tasks.IsEmpty)
                return "";
            
            var contents = new StringBuilder();

            contents.Append("<html><head>");
            contents.Append("<title>todotxt.net</title>");
            contents.Append("<style>" + Resource.CSS + "</style>");
            contents.Append("</head>");

            contents.Append("<body>");
            contents.Append("<h2>todotxt.net</h2>");
            contents.Append("<table>");
            contents.Append("<tr class='tbhead'><th>&nbsp;</th><th>Done</th><th>Created</th><td>Details</td></tr>");

            foreach (Task task in tasks)
            {
                if (task.Completed)
                {
                    contents.Append("<tr class='completedTask'>");
                    contents.Append("<td class='complete'>x</td> ");
                    contents.Append("<td class='completeddate'>" + task.CompletedDate + "</td> ");
                }
                else
                {
                    contents.Append("<tr class='uncompletedTask'>");
                    if (string.IsNullOrEmpty(task.Priority))
                        contents.Append("<td>&nbsp;</td>");
                    else
                        contents.Append("<td><span class='priority'>" + task.Priority + "</span></td>");

                    contents.Append("<td>&nbsp;</td>");
                }

                if (string.IsNullOrEmpty(task.CreationDate))
                    contents.Append("<td>&nbsp;</td>");
                else
                    contents.Append("<td class='startdate'>" + task.CreationDate + "</td>");

                contents.Append("<td>" + task.Body);

                task.Projects.ForEach(project => contents.Append(" <span class='project'>" + project + "</span> "));

                task.Contexts.ForEach(context => contents.Append(" <span class='context'>" + context + "</span> "));

                contents.Append("</td>");

                contents.Append("</tr>");
            }

            contents.Append("</table></body></html>");

            return contents.ToString();
        }
    }
}
