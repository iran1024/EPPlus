/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Constants;
#if !NET35 && !NET45
using System.Collections.Concurrent;
#endif

namespace OfficeOpenXml.Packaging
{
    /// <summary>
    /// Represent an OOXML Zip package.
    /// </summary>
    internal partial class ZipPackage : ZipPackagePartBase, IDisposable
    {
        internal class ContentType
        {
            internal string Name;
            internal bool IsExtension;
            internal string Match;
            public ContentType(string name, bool isExtension, string match)
            {
                Name = name;
                IsExtension = isExtension;
                Match = match;
            }
        }
#if NET35 || NET45
        Dictionary<string, ZipPackagePart> Parts = new Dictionary<string, ZipPackagePart>(StringComparer.OrdinalIgnoreCase);
        internal Dictionary<string, ContentType> _contentTypes = new Dictionary<string, ContentType>(StringComparer.OrdinalIgnoreCase);
#else
        ConcurrentDictionary<string, ZipPackagePart> Parts = new ConcurrentDictionary<string, ZipPackagePart>(StringComparer.OrdinalIgnoreCase);
        internal ConcurrentDictionary<string, ContentType> _contentTypes = new ConcurrentDictionary<string, ContentType>(StringComparer.OrdinalIgnoreCase);
#endif
        internal char _dirSeparator='0';
        internal ZipPackage()
        {
            AddNew();
        }

        private void AddNew()
        {
#if NET35 || NET45
            _contentTypes.Add("xml", new ContentType(ExcelPackage.schemaXmlExtension, true, "xml"));
            _contentTypes.Add("rels", new ContentType(ExcelPackage.schemaRelsExtension, true, "rels"));
#else
            _contentTypes.TryAdd("xml", new ContentType(ExcelPackage.schemaXmlExtension, true, "xml"));
            _contentTypes.TryAdd("rels", new ContentType(ExcelPackage.schemaRelsExtension, true, "rels"));
#endif

        }
        internal ZipInputStream _zip;
        internal ZipPackage(Stream stream)
        {
            bool hasContentTypeXml = false;
            if (stream == null || stream.Length == 0)
            {
                AddNew();
            }
            else
            {
                var rels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                stream.Seek(0, SeekOrigin.Begin);
                //using (ZipInputStream zip = new ZipInputStream(stream))
                //{
                    _zip = new ZipInputStream(stream);
                    var e = _zip.GetNextEntry();
                    if (e == null)
                    {
                        throw (new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor."));
                    }

                    while (e != null)
                    {
                        GetDirSeparator(e);
                        
                        if (e.UncompressedSize > 0)
                        {
                            if (e.FileName.Equals("[content_types].xml", StringComparison.OrdinalIgnoreCase))
                            {
                                AddContentTypes(Encoding.UTF8.GetString(GetZipEntryAsByteArray(_zip, e)));
                                hasContentTypeXml = true;
                            }

                            else if (e.FileName.Equals($"_rels{_dirSeparator}.rels", StringComparison.OrdinalIgnoreCase))
                            {
                                ReadRelation(Encoding.UTF8.GetString(GetZipEntryAsByteArray(_zip, e)), "");
                            }
                            else if (e.FileName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                            {
                                var ba = GetZipEntryAsByteArray(_zip, e);
                                rels.Add(GetUriKey(e.FileName), Encoding.UTF8.GetString(ba));
                            }
                            else
                            {
                                ExtractEntryToPart(_zip, e);
                            }                            
                        }
                        e = _zip.GetNextEntry();
                    }
                    if (_dirSeparator == '0') _dirSeparator = '/';
                    foreach (var p in Parts)
                    {
                        string name = Path.GetFileName(p.Key);
                        string extension = Path.GetExtension(p.Key);
                        string relFile = string.Format("{0}_rels/{1}.rels", p.Key.Substring(0, p.Key.Length - name.Length), name);
                        if (rels.ContainsKey(relFile))
                        {
                            p.Value.ReadRelation(rels[relFile], p.Value.Uri.OriginalString);
                        }
                        if (_contentTypes.ContainsKey(p.Key))
                        {
                            p.Value.ContentType = _contentTypes[p.Key].Name;
                        }
                        else if (extension.Length > 1 && _contentTypes.ContainsKey(extension.Substring(1)))
                        {
                            p.Value.ContentType = _contentTypes[extension.Substring(1)].Name;
                        }
                    }
                    if (!hasContentTypeXml)
                    {
                        throw (new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor."));
                    }
                    if (!hasContentTypeXml)
                    {
                        throw (new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor."));
                    }
                    //zip.Close();
                    //zip.Dispose();
                //}
            }
        }

        private static byte[] GetZipEntryAsByteArray(ZipInputStream zip, ZipEntry e)
        {
            var b = new byte[e.UncompressedSize];
            var size = zip.Read(b, 0, (int)e.UncompressedSize);
            return b;
        }

        private void ExtractEntryToPart(ZipInputStream zip, ZipEntry e)
        {

            var part = new ZipPackagePart(this, e);

            var rest = e.UncompressedSize;
            if (rest > int.MaxValue)
            {
                part.Stream = null; //Over 2GB, we use the zip stream directly instead.
            }
            else
            {
                const int BATCH_SIZE = 0x100000;
                part.Stream = RecyclableMemory.GetStream();
                while (rest > 0)
                {
                    var bufferSize = rest > BATCH_SIZE ? BATCH_SIZE : (int)rest;
                    var b = new byte[bufferSize];
                    var size = zip.Read(b, 0, bufferSize);
                    part.Stream.Write(b, 0, size);
                    rest -= size;
                }
            }
#if NET35 || NET45
            Parts.Add(GetUriKey(e.FileName), part);
#else
            Parts.TryAdd(GetUriKey(e.FileName), part);
#endif
        }

        private void GetDirSeparator(ZipEntry e)
        {
            if (_dirSeparator == '0')
            {
                if (e.FileName.Contains("\\"))
                {
                    _dirSeparator = '\\';
                }
                else if (e.FileName.Contains("/"))
                {
                    _dirSeparator = '/';
                }
            }
        }

        private void AddContentTypes(string xml)
        {
            var doc = new XmlDocument();
            XmlHelper.LoadXmlSafe(doc, xml, Encoding.UTF8);

            foreach (XmlElement c in doc.DocumentElement.ChildNodes)
            {
                ContentType ct;
                if (string.IsNullOrEmpty(c.GetAttribute("Extension")))
                {
                    ct = new ContentType(c.GetAttribute("ContentType"), false, c.GetAttribute("PartName"));
#if NET35 || NET45
                    _contentTypes.Add(GetUriKey(ct.Match), ct);
#else
                    _contentTypes.TryAdd(GetUriKey(ct.Match), ct);
#endif
                }
                else
                {
                    ct = new ContentType(c.GetAttribute("ContentType"), true, c.GetAttribute("Extension"));
#if NET35 || NET45
                    _contentTypes.Add(ct.Match, ct);
#else
                    _contentTypes.TryAdd(ct.Match, ct);
#endif
                }
            }
        }

#region Methods
        internal ZipPackagePart CreatePart(Uri partUri, string contentType)
        {
            return CreatePart(partUri, contentType, CompressionLevel.Default);
        }
        internal ZipPackagePart CreatePart(Uri partUri, string contentType, CompressionLevel compressionLevel, string extension=null)
        {
            if (PartExists(partUri))
            {
                throw (new InvalidOperationException("Part already exist"));
            }

            var part = new ZipPackagePart(this, partUri, contentType, compressionLevel);
            if(string.IsNullOrEmpty(extension))
            {
#if NET35 || NET45
                _contentTypes.Add(GetUriKey(part.Uri.OriginalString), new ContentType(contentType, false, part.Uri.OriginalString));
#else
                _contentTypes.TryAdd(GetUriKey(part.Uri.OriginalString), new ContentType(contentType, false, part.Uri.OriginalString));
#endif
            }
            else
            {
                if (!_contentTypes.ContainsKey(extension))
                {
#if NET35 || NET45
                    _contentTypes.Add(extension, new ContentType(contentType, true, extension));
#else
                    _contentTypes.TryAdd(extension, new ContentType(contentType, true, extension));
#endif
                }
            }
#if NET35 || NET45
            Parts.Add(GetUriKey(part.Uri.OriginalString), part);
#else
            Parts.TryAdd(GetUriKey(part.Uri.OriginalString), part);
#endif
            return part;
        }
        internal ZipPackagePart CreatePart(Uri partUri, ZipPackagePart sourcePart)
        {
            var destPart = CreatePart(partUri, sourcePart.ContentType);
            var destStream = destPart.GetStream(FileMode.Create, FileAccess.Write);
            var sourceStream = sourcePart.GetStream();
            var b = ((MemoryStream)sourceStream).GetBuffer();
            destStream.Write(b, 0, b.Length);
            destStream.Flush();
            return destPart;
        }
        internal ZipPackagePart CreatePart(Uri partUri, string contentType, string xml)
        {
            var destPart = CreatePart(partUri, contentType);
            var destStream = new StreamWriter(destPart.GetStream(FileMode.Create, FileAccess.Write));
            destStream.Write(xml);
            destStream.Flush();
            return destPart;
        }

        internal ZipPackagePart GetPart(Uri partUri)
        {
            if (PartExists(partUri))
            {
                return Parts.Single(x => x.Key.Equals(GetUriKey(partUri.OriginalString),StringComparison.OrdinalIgnoreCase)).Value;
            }
            else
            {
                throw (new InvalidOperationException("Part does not exist."));
            }
        }

        internal string GetUriKey(string uri)
        {
            string ret = uri.Replace('\\', '/');
            if (ret[0] != '/')
            {
                ret = '/' + ret;
            }
            return ret;
        }
        internal bool PartExists(Uri partUri)
        {
            var uriKey = GetUriKey(partUri.OriginalString.ToLowerInvariant());
            return Parts.ContainsKey(uriKey);
        }
#endregion

        internal void DeletePart(Uri Uri)
        {
            var delList=new List<object[]>(); 
            foreach (var p in Parts.Values)
            {
                foreach (var r in p.GetRelationships())
                {                    
                    if (r.TargetUri !=null && UriHelper.ResolvePartUri(p.Uri, r.TargetUri).OriginalString.Equals(Uri.OriginalString, StringComparison.OrdinalIgnoreCase))
                    {                        
                        delList.Add(new object[]{r.Id, p});
                    }
                }
            }
            foreach (var o in delList)
            {
                ((ZipPackagePart)o[1]).DeleteRelationship(o[0].ToString());
            }
            var rels = GetPart(Uri).GetRelationships();
            while (rels.Count > 0)
            {
                rels.Remove(rels.First().Id);
            }
            rels=null;

#if NET35 || NET45
            _contentTypes.Remove(GetUriKey(Uri.OriginalString));
            //remove all relations
            Parts.Remove(GetUriKey(Uri.OriginalString));
#else
            _contentTypes.TryRemove(GetUriKey(Uri.OriginalString), out _);
            Parts.TryRemove(GetUriKey(Uri.OriginalString), out _);
#endif
        }
        internal void Save(Stream stream)
        {
            var enc = Encoding.UTF8;
            ZipOutputStream os = new ZipOutputStream(stream, true);
            os.EnableZip64 = Zip64Option.AsNecessary;
            os.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)_compression;            

            /**** ContentType****/
            var entry = os.PutNextEntry("[Content_Types].xml");
            byte[] b = enc.GetBytes(GetContentTypeXml());
            os.Write(b, 0, b.Length);
            /**** Top Rels ****/
            _rels.WriteZip(os, $"_rels/.rels");
            ZipPackagePart ssPart=null;

            foreach(var part in Parts.Values)
            {
                if (part.ContentType != ContentTypes.contentTypeSharedString)
                {
                    part.WriteZip(os);
                }
                else
                {
                    ssPart = part;
                }
            }

            //Shared strings must be saved after all worksheets. The ss dictionary is populated when that workheets are saved (to get the best performance).
            if (ssPart != null)
            {
                ssPart.WriteZip(os);
            }
            os.Flush();
            
            os.Close();
            os.Dispose();  
            
            //return ms;
        }

        private string GetContentTypeXml()
        {
            StringBuilder xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            foreach (ContentType ct in _contentTypes.Values)
            {
                if (ct.IsExtension)
                {
                    xml.AppendFormat("<Default ContentType=\"{0}\" Extension=\"{1}\"/>", ct.Name, ct.Match);
                }
                else
                {
                    xml.AppendFormat("<Override ContentType=\"{0}\" PartName=\"{1}\" />", ct.Name, GetUriKey(ct.Match));
                }
            }
            xml.Append("</Types>");
            return xml.ToString();
        }
        internal void Flush()
        {

        }
        internal void Close()
        {
            
        }

        public void Dispose()
        {
            foreach(var part in Parts.Values)
            {
                part.Dispose();
            }
            _zip?.Dispose();
        }

        CompressionLevel _compression = CompressionLevel.Default;
        /// <summary>
        /// Compression level
        /// </summary>
        public CompressionLevel Compression 
        { 
            get
            {
                return _compression;
            }
            set
            {
                foreach (var part in Parts.Values)
                {
                    if (part.CompressionLevel == _compression)
                    {
                        part.CompressionLevel = value;
                    }
                }
                _compression = value;
            }
        }
    }
}
