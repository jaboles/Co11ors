using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.IO;
using Vestris.ResourceLib;

namespace VSIconSwitcher
{
    public class Range
    {
        public Range(int id)
        {
            Lower = id;
            Upper = id;
        }
        public Range(int lower, int upper)
        {
            Lower = lower;
            Upper = upper;
            Debug.Assert(lower <= upper, "Lower should be lower than Upper end of range");
        }
        public int Lower { get; private set; }
        public int Upper { get; private set; }
    }


    public class NativeResourceReplacement : AssetReplacement
    {
        public static void Initialize(Func<string, string> pathTranslator)
        {
            s_pathTranslate = pathTranslator;
        }

        public NativeResourceReplacement(string filename, ResourceType resType)
            : base(filename)
        {
            m_resType = resType;
        }

        public NativeResourceReplacement(string filename, ResourceType resType, params string[] ids)
            : this(filename, resType)
        {
            m_ids = ids;
        }
        public NativeResourceReplacement(string filename, ResourceType resType, string id)
            : this(filename, resType)
        {
            m_ids = new string[] { id };
        }
        public NativeResourceReplacement(string filename, ResourceType resType, params Range[] ranges)
            : this(filename, resType)
        {
            List<string> ids = new List<string>();
            foreach (Range range in ranges)
            {
                for (int i = range.Lower; i <= range.Upper; i++)
                    ids.Add(i.ToString());
            }
            m_ids = ids.ToArray();
        }

        public override void CopyResources(string src, string dest)
        {
            ResourceInfo srcInfo = new ResourceInfo();
            ResourceInfo dstInfo = new ResourceInfo();
            srcInfo.Load(src);
            dstInfo.Load(dest);

            Kernel32.ResourceTypes resType = 0;
            switch (ResourceType)
            {
                case ResourceType.Icon:
                    resType = Kernel32.ResourceTypes.RT_GROUP_ICON;
                    break;
                case ResourceType.Bitmap:
                    resType = Kernel32.ResourceTypes.RT_BITMAP;
                    break;
                default:
                    Debug.Fail("Unknown resource type");
                    break;
            }

            List<Resource> resourcesToReplace = new List<Resource>();
            foreach (Resource res in srcInfo[resType])
            {
                if (Ids == null || Ids.Contains(res.Name.Name))
                {
                    // Find matching resource in destination file
                    Resource dstRes = dstInfo[resType].Single(dr => dr.Name.Equals(res.Name));

                    res.LoadFrom(src);
                    dstRes.LoadFrom(dest);

                    res.Language = dstRes.Language;

                    // Do some checks
                    if (ResourceType == VSIconSwitcher.ResourceType.Icon)
                    {
                        IconDirectoryResource srcIconDirectory = res as IconDirectoryResource;
                        IconDirectoryResource dstIconDirectory = dstRes as IconDirectoryResource;
                        //Debug.Assert(srcIconDirectory.Icons.Count == dstIconDirectory.Icons.Count, "Source/destination icon directory count should match.");
                    }
                    else if (ResourceType == VSIconSwitcher.ResourceType.Bitmap)
                    {
                        BitmapResource srcBitmap = res as BitmapResource;
                        BitmapResource dstBitmap = dstRes as BitmapResource;
                        Debug.Assert(srcBitmap.Bitmap.Header.biHeight == dstBitmap.Bitmap.Header.biHeight, "Source/destination bitmap height dimension should match.");
                        Debug.Assert(srcBitmap.Bitmap.Header.biWidth == dstBitmap.Bitmap.Header.biWidth, "Source/destination bitmap height dimension should match.");
                    }

                    resourcesToReplace.Add(res);
                }
            }

            dstInfo.Unload();

            foreach (Resource res in resourcesToReplace)
            {
                res.SaveTo(dest);
            }

            srcInfo.Unload();
        }

        public IEnumerable<string> Ids { get { return m_ids; } }
        public bool NeedsGAC
        {
            get
            {
                if (!m_needsGac.HasValue)
                    m_needsGac = NetFXUtilities.IsInGAC(FilePath);
                Debug.Assert(m_needsGac.HasValue);
                return m_needsGac.Value;
            }
        }
        public ResourceType ResourceType { get { return m_resType; } }

        private IEnumerable<string> m_ids;
        private bool? m_needsGac;
        private ResourceType m_resType;
        private static Func<string, string> s_pathTranslate;
    }
}
