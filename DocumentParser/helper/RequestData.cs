using System;
using System.Collections.Generic;

namespace DocumentParser.helper
{
    [Serializable]
    public class RequestData
    {
        private string serviceName;

        public string ServiceName
        {
            get { return serviceName; }
            set { serviceName = value; }
        }

        private int sequence;

        public int Sequence
        {
            get { return sequence; }
            set { sequence = value; }
        }

        private string docName;

        public string DocName
        {
            get { return docName; }
            set { docName = value; }
        }

        private string stringParam;

        public string StringParam
        {
            get { return stringParam; }
            set { stringParam = value; }
        }

        private string filename;

        public string Filename
        {
            get { return filename; }
            set { filename = value; }
        }

        private int intParam;

        private int indent;

        public int Indent
        {
            get { return indent; }
            set { indent = value; }
        }

        public int IntParam
        {
            get { return intParam; }
            set { intParam = value; }
        }
       
        private string splitParam;

        public string SplitParam
        {
            get { return splitParam; }
            set { splitParam = value; }
        }

    }
}
