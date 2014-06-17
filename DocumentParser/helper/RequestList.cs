using System.Collections.Generic;

namespace DocumentParser.helper
{
    public class RequestList
    {
        private List<RequestData> requestData = new List<RequestData>();

        public List<RequestData> RequestData
        {
            get { return requestData; }
            set { requestData = value; }
        }

    }
}
