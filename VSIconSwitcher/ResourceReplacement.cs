using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace VSIconSwitcher
{
    public class ResourceReplacement
    {
        public ResourceReplacement(string filename, ResourceType resType, int id)
        {
            m_sourceFile = filename;
            m_destinationFile = filename;
            m_resType = resType;
            m_sourceId = id;
            m_destinationId = id;
        }

        public string SourceFile { get { return m_sourceFile; } }
        public int SourceId { get { return m_sourceId; } }
        public string DestinationFile { get { return m_destinationFile; } }
        public int DestinationId { get { return m_destinationId; } }
        public bool NeedsGAC
        {
            get
            {
                if (!m_needsGac.HasValue)
                    m_needsGac = NetFXUtilities.IsInGAC(SourceFile);
                Debug.Assert(m_needsGac.HasValue);
                return m_needsGac.Value;
            }
        }
        public ResourceType ResourceType { get { return m_resType; } }

        private string m_sourceFile;
        private int m_sourceId;
        private string m_destinationFile;
        private int m_destinationId;
        private bool? m_needsGac;
        private ResourceType m_resType;
    }
}
