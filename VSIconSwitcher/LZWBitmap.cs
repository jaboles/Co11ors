using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace VSIconSwitcher
{
    public class LZWBitmap
    {
        private int m_width;
        private int m_height;
        private int m_type;

        public LZWBitmap()
        {
        }

        public void Load(MemoryStream ms)
        {
            m_width = ms.ReadByte() + (ms.ReadByte() << 8);
            m_height = ms.ReadByte() + (ms.ReadByte() << 8);
            m_height = ms.ReadByte() + (ms.ReadByte() << 8);
        }
    }
}
