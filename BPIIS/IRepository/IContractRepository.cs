using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApp1.IRepository
{
    public interface IContractRepository
    {
        string GetNo(string wholeText);
        string GetName(string wholeText);

        string GetAmount(string wholeText);

        string GetProjectLocation(string wholeText);

        string GetSignedDate(string wholeText);

        string GetJobContent(string wholeText);

        string GetClient(string wholeText);

        string GetClientContactPerson(string wholeText);

        string GetClientContactPersonPhone(string wholeText);

        string GetDeadline(string wholeText);
    }
}
