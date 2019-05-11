﻿// Copyright (c) 2016, Dijji, and released under Ms-PL.  This can be found in the root of this distribution. 

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Text.RegularExpressions;

namespace XstReader
{
    // Main handling for xst (.ost and .pst) files
    //
    // The code here implements the messaging layer, which depends on and invokes the NDP and LTP layers
    //
    // The constructor provides the path to the file, and then there are just a few public methods:
    // - Read the structure of the folders
    // - Read the list of messages contained in a folder
    // - Read the contents of a message
    // - Save an attachment to a message

    class XstFile
    {
        private NDB ndb;
        private LTP ltp;
        private View view;


        // We use sets of PropertyGetters to define the equivalent of queries when reading property sets and tables

        // The folder properties we read when exploring folder structure
        private static readonly PropertyGetters<Folder> pgFolder = new PropertyGetters<Folder>
        {
            {EpropertyTag.PidTagDisplayName, (f, val) => f.Name = val },
            {EpropertyTag.PidTagContentCount, (f, val) => f.ContentCount = val },
            // Don't bother reading HasSubFolders, because it is not always set
            // {EpropertyTag.PidTagSubfolders, (f, val) => f.HasSubFolders = val },
        };

        // When reading folder contents, the message properties we ask for
        private static readonly PropertyGetters<Message> pgMessageList = new PropertyGetters<Message>
        {
            {EpropertyTag.PidTagSubjectW, (m, val) => m.Subject = val },
            {EpropertyTag.PidTagDisplayCcW, (m, val) => m.Cc = val },
            {EpropertyTag.PidTagDisplayToW, (m, val) => m.To = val },
            {EpropertyTag.PidTagMessageFlags, (m, val) => m.Flags = (MessageFlags)val },
            {EpropertyTag.PidTagSentRepresentingNameW, (m, val) => m.From = val },
            {EpropertyTag.PidTagClientSubmitTime, (m, val) => m.Submitted = val },
            {EpropertyTag.PidTagMessageDeliveryTime, (m, val) => m.Received = val },
            {EpropertyTag.PidTagLastModificationTime, (m, val) => m.Modified = val },
        };

        // When reading folder contents, the message properties we ask for
        // In Unicode4K, PidTagSentRepresentingNameW doesn't yield a useful value
        private static readonly PropertyGetters<Message> pgMessageList4K = new PropertyGetters<Message>
        {
            {EpropertyTag.PidTagSubjectW, (m, val) => m.Subject = val },
            {EpropertyTag.PidTagDisplayCcW, (m, val) => m.Cc = val },
            {EpropertyTag.PidTagDisplayToW, (m, val) => m.To = val },
            {EpropertyTag.PidTagMessageFlags, (m, val) => m.Flags = (MessageFlags)val },
            {EpropertyTag.PidTagClientSubmitTime, (m, val) => m.Submitted = val },
            {EpropertyTag.PidTagMessageDeliveryTime, (m, val) => m.Received = val },
            {EpropertyTag.PidTagLastModificationTime, (m, val) => m.Modified = val },
        };

        private static readonly PropertyGetters<Message> pgMessageDetail4K = new PropertyGetters<Message>
        {
            {EpropertyTag.PidTagSentRepresentingNameW, (m, val) => m.From = val },
            {EpropertyTag.PidTagSentRepresentingEmailAddress, (m, val) => { if(m.From == null) m.From = val; } },
            {EpropertyTag.PidTagSenderName, (m, val) => { if(m.From == null) m.From = val; } },
        };

        // The properties we read when accessing the contents of a message
        private static readonly PropertyGetters<Message> pgMessageContent = new PropertyGetters<Message>
        {
            {EpropertyTag.PidTagNativeBody, (m, val) => m.NativeBody = (BodyType)val },
            {EpropertyTag.PidTagBody, (m, val) => m.Body = val },
            //{EpropertyTag.PidTagInternetCodepage, (m, val) => m.InternetCodePage = (int)val },
            // In ANSI format, PidTagHtml is called PidTagBodyHtml (though the tag code is the same), because it is a string rather than a binary value
            // Here, we test the type to determine where to put the value 
            {EpropertyTag.PidTagHtml, (m, val) => { if (val is string)  m.BodyHtml = val; else m.Html = val; } },
            {EpropertyTag.PidTagRtfCompressed, (m, val) => m.RtfCompressed = val },
        };

        // The properties we read when accessing the recipient table of a message
        private static readonly PropertyGetters<Recipient> pgMessageRecipient = new PropertyGetters<Recipient>
        {
            {EpropertyTag.PidTagRecipientType, (r, val) => r.RecipientType = (RecipientType)val },
            {EpropertyTag.PidTagDisplayName, (r, val) => r.DisplayName = val },
            {EpropertyTag.PidTagEmailAddress, (r, val) => r.EmailAddress = val },
        };

        //The properties we read when accessing a message attached to a message
        private static readonly PropertyGetters<Message> pgMessageAttachment = new PropertyGetters<Message>
        {
            {EpropertyTag.PidTagSubjectW, (m, val) => m.Subject = val },
            {EpropertyTag.PidTagDisplayCcW, (m, val) => m.Cc = val },
            {EpropertyTag.PidTagDisplayToW, (m, val) => m.To = val },
            {EpropertyTag.PidTagMessageFlags, (m, val) => m.Flags = (MessageFlags)val },
            {EpropertyTag.PidTagSentRepresentingNameW, (m, val) => m.From = val },
            {EpropertyTag.PidTagClientSubmitTime, (m, val) => m.Submitted = val },
            {EpropertyTag.PidTagMessageDeliveryTime, (m, val) => m.Received = val },
            {EpropertyTag.PidTagLastModificationTime, (m, val) => m.Modified = val },
            {EpropertyTag.PidTagNativeBody, (m, val) => m.NativeBody = (BodyType)val },
            {EpropertyTag.PidTagBody, (m, val) => m.Body = val },
            {EpropertyTag.PidTagHtml, (m, val) => { if (val is string)  m.BodyHtml = val; else m.Html = val; } },
            {EpropertyTag.PidTagRtfCompressed, (m, val) => m.RtfCompressed = val },
        };

        private static readonly HashSet<EpropertyTag> contentExclusions = new HashSet<EpropertyTag>
        {
            EpropertyTag.PidTagNativeBody,
            EpropertyTag.PidTagBody,
            EpropertyTag.PidTagHtml,
            EpropertyTag.PidTagRtfCompressed,
        };

        // The properties we read when getting a list of attachments
        private static readonly PropertyGetters<Attachment> pgAttachmentList = new PropertyGetters<Attachment>
        {
            {EpropertyTag.PidTagDisplayName, (a, val) => a.DisplayName = val },
            {EpropertyTag.PidTagAttachFilenameW, (a, val) => a.FileNameW = val },
            {EpropertyTag.PidTagAttachLongFilename, (a, val) => a.LongFileName = val },
            {EpropertyTag.PidTagAttachmentSize, (a, val) => a.Size = val },
            {EpropertyTag.PidTagAttachMethod, (a, val) => a.AttachMethod = (AttachMethods)val },
            //{EpropertyTag.PidTagAttachMimeTag, (a, val) => a.MimeTag = val },
            {EpropertyTag.PidTagAttachPayloadClass, (a, val) => a.FileNameW = val },
        };

        // The properties we read To enable handling of HTML images delivered as attachments
        private static readonly PropertyGetters<Attachment> pgAttachedHtmlImages = new PropertyGetters<Attachment>
        {
            {EpropertyTag.PidTagAttachFlags, (a, val) => a.Flags = (AttachFlags)val },
            {EpropertyTag.PidTagAttachMimeTag, (a, val) => a.MimeTag = val },
            {EpropertyTag.PidTagAttachContentId, (a, val) => a.ContentId = val },
            {EpropertyTag.PidTagAttachmentHidden, (a, val) => a.Hidden = val },
        };

        // The properties we read when accessing the name of an attachment
        private static readonly PropertyGetters<Attachment> pgAttachmentName = new PropertyGetters<Attachment>
        {
            {EpropertyTag.PidTagAttachFilenameW, (a, val) => a.FileNameW = val },
            {EpropertyTag.PidTagAttachLongFilename, (a, val) => a.LongFileName = val },
        };

        // The properties we read when accessing the contents of an attachment
        private static readonly PropertyGetters<Attachment> pgAttachmentContent = new PropertyGetters<Attachment>
        {
            {EpropertyTag.PidTagAttachDataBinary, (a, val) => a.Content = val },
        };

        private static readonly HashSet<EpropertyTag> attachmentContentExclusions = new HashSet<EpropertyTag>
        {
            EpropertyTag.PidTagAttachDataBinary,
        };


        #region Public methods

        public XstFile(View view, string fullName)
        {
            this.ndb = new NDB(fullName);
            this.ltp = new LTP(ndb);
            this.view = view;
        }

        public void ReadFolderTree()
        {
            ndb.Initialise();

            using (var fs = ndb.GetReadStream())
            {
                var root = ReadFolderStructure(fs, new NID(EnidSpecial.NID_ROOT_FOLDER));

                foreach (var f in root.Folders)
                {
                    // We may be called on a background thread, so we need to dispatch this to the UI thread
                    Application.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        view.RootFolders.Add(f);
                    }));
                }
            }
        }

        public void ReadMessages(Folder f)
        {
            if (f.ContentCount > 0)
            {
                using (var fs = ndb.GetReadStream())
                {
                    // Get the Contents table for the folder
                    // For 4K, not all the properties we want are available in the Contents table, so supplement them from the Message itself
                    var ms = ltp.ReadTable<Message>(fs, NID.TypedNID(EnidType.CONTENTS_TABLE, f.Nid),
                                                    ndb.IsUnicode4K ? pgMessageList4K : pgMessageList, (m, id) => m.Nid = new NID(id))
                                .Select(m => ndb.IsUnicode4K ? Add4KMessageProperties(fs, m) : m);

                    // We may be called on a background thread, so we need to dispatch this to the UI thread
                    Application.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        f.Messages.Clear();
                        foreach (var m in ms)
                        {
                            f.AddMessage(m);
                        }
                    }));
                }
            }
        }

        public void ReadMessageDetails(Message m)
        {
            using (var fs = ndb.GetReadStream())
            {
                // Read the contents properties
                var subNodeTree = ltp.ReadProperties<Message>(fs, m.Nid, pgMessageContent, m);

                // Read all other properties
                m.Properties.Clear();
                foreach (var p in ltp.ReadAllProperties(fs, m.Nid, contentExclusions))
                {
                    m.Properties.Add(p);
                }

                ReadMessageTables(fs, subNodeTree, m);
            }
        }

        public List<Property> ReadAttachmentProperties(Attachment a)
        {
            using (var fs = ndb.GetReadStream())
            {
                BTree<Node> subNodeTreeMessage = a.subNodeTreeProperties;

                if (subNodeTreeMessage == null)
                    // No subNodeTree given: assume we can look it up in the main tree
                    ndb.LookupNodeAndReadItsSubNodeBtree(fs, a.Parent.Nid, out subNodeTreeMessage);

                // Read all non-content properties
                // Convert to list so that we can dispose the file access
                return new List<Property>(ltp.ReadAllProperties(fs, subNodeTreeMessage, a.Nid, attachmentContentExclusions, true));
            }
        }

        public void SaveAttachment(string path, Attachment a)
        {
            using (var afs = new FileStream(Path.Combine(path, a.FileName), FileMode.Create, FileAccess.Write))
            {
                SaveAttachment(afs, a);
            }
        }

        public void SaveAttachment(Stream s, Attachment a)
        {
            using (FileStream fs = ndb.GetReadStream())
            {
                BTree<Node> subNodeTreeMessage = a.subNodeTreeProperties;

                if (subNodeTreeMessage == null)
                    // No subNodeTree given: assume we can look it up in the main tree
                    ndb.LookupNodeAndReadItsSubNodeBtree(fs, a.Parent.Nid, out subNodeTreeMessage);

                var subNodeTreeAttachment = ltp.ReadProperties<Attachment>(fs, subNodeTreeMessage, a.Nid, pgAttachmentContent, a);

                // If the value is inline, we just write it out
                if (a.Content.GetType() == typeof(byte[]))
                {
                    s.Write(a.Content, 0, a.Content.Length);
                }
                // Otherwise we need to dereference the node pointing to the data,
                // using the subnode tree belonging to the attachment
                else if (a.Content.GetType() == typeof(NID))
                {
                    var nb = NDB.LookupSubNode(subNodeTreeAttachment, (NID)a.Content);

                    // Copy the data to the output file stream without getting it all into memory at once,
                    // as there can be a lot of data
                    ndb.CopyDataBlocks(fs, s, nb.DataBid);
                }
            }
        }

        public Message OpenAttachedMessage(Attachment a)
        {
            using (FileStream fs = ndb.GetReadStream())
            {
                BTree<Node> subNodeTreeMessage = a.subNodeTreeProperties;

                if (subNodeTreeMessage == null)
                    // No subNodeTree given: assume we can look it up in the main tree
                    ndb.LookupNodeAndReadItsSubNodeBtree(fs, a.Parent.Nid, out subNodeTreeMessage);

                var subNodeTreeAttachment = ltp.ReadProperties<Attachment>(fs, subNodeTreeMessage, a.Nid, pgAttachmentContent, a);

                if (a.Content.GetType() == typeof(PtypObjectValue))
                {
                    Message m = new Message { Nid = new NID(((PtypObjectValue)a.Content).Nid) };

                    // Read the basic and contents properties
                    var childSubNodeTree = ltp.ReadProperties<Message>(fs, subNodeTreeAttachment, m.Nid, pgMessageAttachment, m, true);

                    // Read all other properties
                    foreach (var p in ltp.ReadAllProperties(fs, subNodeTreeAttachment, m.Nid, contentExclusions, true))
                    {
                        m.Properties.Add(p);
                    }

                    ReadMessageTables(fs, childSubNodeTree, m, true);

                    return m;
                }
                else
                    throw new XstException("Unexpected data type for attached message");
            }
        }

        private struct LineProp
        {
            public int line;
            public Property p;
        }

        public void ExportMessageProperties(IEnumerable<Message> messages, string fileName)
        {
            // We build a dictionary of queues of line,Property pairs where each queue represents
            // a column in the CSV file, and the line is the line number in the file.
            // The key to the dictionary is the property ID.

            var dict = new Dictionary<string, Queue<LineProp>>();
            int lines = 1;

            Queue<LineProp> queue_body = new Queue<LineProp>();
            dict["0000Body"] = queue_body;
            Queue<LineProp> queue_html_utf8 = new Queue<LineProp>();
            dict["0000HTML_utf8"] = queue_html_utf8;
            Queue<LineProp> queue_html_base64 = new Queue<LineProp>();
            dict["0000HTML_base64"] = queue_html_base64;
            Queue<LineProp> queue_rtf = new Queue<LineProp>();
            dict["0000RTF"] = queue_rtf;
            
            foreach (var m in messages)
            {
                // Do not reread properties for current message as it will fail updating the display
                if (m != view.CurrentMessage)
                {
                    try
                    {
                        ReadMessageDetails(m);
                    }
                    catch (XstException ex)
                    {
                        // Ignore file exceptions to get as much as we can
                    }
                }


                foreach (var p in m.Properties)
                {
                    Queue<LineProp> queue;
                    if (!dict.TryGetValue(p.CsvId, out queue))
                    {
                        queue = new Queue<LineProp>();
                        dict[p.CsvId] = queue;
                    }
                    queue.Enqueue(new LineProp { line = lines, p = p });
                }

                if ( m.Body != null )
                {
                    queue_body.Enqueue(new LineProp { line = lines, p = new Property { Tag = EpropertyTag.PidTagBody, Value = m.Body } });
                }

                if (m.Html != null)
                {
                    var body_html_utf8string = System.Text.Encoding.UTF8.GetString(m.Html);
                    // queue_html_default.Enqueue(new LineProp { line = lines, p = new Property { Tag = EpropertyTag.PidTagHtml, Value = System.Text.Encoding.GetEncodingi("sjis") } });
                    // queue_html_default.Enqueue(new LineProp { line = lines, p = new Property { Tag = EpropertyTag.PidTagHtml, Value = System.Text.Encoding.Default.GetString( m.Html) } });
                    // queue_html_unicode.Enqueue(new LineProp { line = lines, p = new Property { Tag = EpropertyTag.PidTagHtml, Value = System.Text.Encoding.Unicode.GetString(m.Html) } });
                    queue_html_utf8   .Enqueue(new LineProp { line = lines, p = new Property { Tag = EpropertyTag.PidTagNativeBody, Value = body_html_utf8string } });  // Tag is misuse but no appropriate one --HAL

                    var charset = "utf-8";
                    var r = new Regex(@"charset=""?([-_a-zA-Z0-9]+)");
                    var match = r.Match(body_html_utf8string);
                    //System.Text.RegularExpressions.Match m = r.Match(text);

                    if (match.Success)
                    {
                        charset = match.Groups[1].Value;
                    }

                    queue_html_base64.Enqueue(new LineProp { line = lines,
                                                           p = new Property { Tag = EpropertyTag.PidTagHtml,
                                                                              Value = "Content-Type: text/html; charset=" + charset + "\nContent-Transfer-Encoding: base64\n\n" + 
                                                                                      Convert.ToBase64String(m.Html, Base64FormattingOptions.InsertLineBreaks) } });
                }

                if (m.RtfCompressed != null)
                {
                    queue_rtf.Enqueue(new LineProp {
                        line = lines,
                        p = new Property {
                            Tag = EpropertyTag.PidTagRtfCompressed, 
                            Value = "Content-Type: application/ms-tnef; name = \"winmail.dat\"\nContent-Transfer-Encoding: base64\n\n" +
                                                                            Convert.ToBase64String(m.RtfCompressed, Base64FormattingOptions.InsertLineBreaks)
                        }
                    });
                }


                lines++;
            }
            if (queue_body.Count == 0)
            {
                queue_body.Enqueue(new LineProp
                {
                    line = 0,
                    p = new Property
                    {
                        Tag = EpropertyTag.PidTagBody,
                        Value = ""
                    }
                });
            }

            if (queue_html_utf8.Count == 0)
            {
                queue_html_utf8.Enqueue(new LineProp
                {
                    line = 0,
                    p = new Property
                    {
                        Tag = EpropertyTag.PidTagNativeBody,
                        Value = ""
                    }
                });
            }

            if (queue_html_base64.Count == 0)
            {
                queue_html_base64.Enqueue(new LineProp
                {
                    line = 0,
                    p = new Property
                    {
                        Tag = EpropertyTag.PidTagHtml,
                        Value = ""
                    }
                });
            }

            if (queue_rtf.Count == 0)
            {
                queue_rtf.Enqueue(new LineProp
                {
                    line = 0,
                    p = new Property
                    {
                        Tag = EpropertyTag.PidTagRtfCompressed,
                        Value = ""
                    }
                });
            }

            // Now we sort the columns by ID
            var columns = dict.Keys.OrderBy(x => x).ToArray();

            // And finally output the CSV file line by line
            using (var sw = new System.IO.StreamWriter(fileName, false,  /* Encoding.Default */ System.Text.Encoding.GetEncoding("utf-8") ))

            {
                StringBuilder sb = new StringBuilder();
                bool hasValue = false;

                for (int line = 0; line < lines; line++)
                {
                    foreach (var col in columns)
                    {
                        var q = dict[col];

                        // First line is always the column headers
                        if (line == 0)
                            AddCsvValue(sb, q.Peek().p.CsvDescription, ref hasValue);

                        // After that, output the column value if it has one
                        else if (q.Count > 0 && q.Peek().line == line)
                            AddCsvValue(sb, q.Dequeue().p.DisplayValue, ref hasValue);
                        
                        // Or leave it blank
                        else
                            AddCsvValue(sb, "", ref hasValue);
                    }

                    // Write the completed line out
                    sw.WriteLine(sb.ToString());
                    sb.Clear();
                    hasValue = false;
                }
            }
        }
        #endregion

        #region Private methods

        // Recurse down the folder tree, building a structure of Folder classes
        private Folder ReadFolderStructure(FileStream fs, NID nid)
        {
            Folder f = new Folder { Nid = nid };

            ltp.ReadProperties<Folder>(fs, nid, pgFolder, f);

            foreach (var sf in ltp.ReadTableRowIds(fs, NID.TypedNID(EnidType.HIERARCHY_TABLE, nid))
                .Where(id => id.nidType == EnidType.NORMAL_FOLDER)
                .Select(id => ReadFolderStructure(fs, id))
                .OrderBy(sf => sf.Name))
                f.Folders.Add(sf);

            return f;
        }

        private Message Add4KMessageProperties(FileStream fs, Message m)
        {
            ltp.ReadProperties<Message>(fs, m.Nid, pgMessageDetail4K, m);

            return m;
        }

        private void ReadMessageTables(FileStream fs, BTree<Node> subNodeTree, Message m, bool isAttached = false)
        {
            // Read the recipient table for the message
            var recipientsNid = new NID(EnidSpecial.NID_RECIPIENT_TABLE);
            if (ltp.IsTablePresent(subNodeTree, recipientsNid))
            {
                var rs = ltp.ReadTable<Recipient>(fs, subNodeTree, recipientsNid, pgMessageRecipient, null, (r, p) => r.Properties.Add(p));
                m.Recipients.Clear();
                foreach (var r in rs)
                {
                    // Sort the properties
                    List<Property> lp = new List<Property>(r.Properties);
                    lp.Sort((a, b) => a.Tag.CompareTo(b.Tag));
                    r.Properties.Clear();
                    foreach (var p in lp)
                        r.Properties.Add(p);

                    m.Recipients.Add(r);
                }
            }

            // Read any attachments
            var attachmentsNid = new NID(EnidSpecial.NID_ATTACHMENT_TABLE);
            if (m.HasAttachment)
            {
                // Read the attachment table, which is held in the subnode of the message
                var atts = ltp.ReadTable<Attachment>(fs, subNodeTree, attachmentsNid, pgAttachmentList, (a, id) => a.Nid = new NID(id)).ToList();
                foreach (var a in atts)
                {
                    a.XstFile = this; // For lazy reading of the complete properties
                    a.Parent = m;

                    // If the long name wasn't in the attachment table, go look for it in the attachment properties
                    if (a.LongFileName == null)
                        ltp.ReadProperties<Attachment>(fs, subNodeTree, a.Nid, pgAttachmentName, a);

                    // Read properties relating to HTML images presented as attachments
                    ltp.ReadProperties<Attachment>(fs, subNodeTree, a.Nid, pgAttachedHtmlImages, a);

                    // If this is an embedded email, tell the attachment where to look for its properties
                    // This is needed because the email node is not in the main node tree
                    if (isAttached)
                        a.subNodeTreeProperties = subNodeTree;
                }

                m.SortAndSaveAttachments(atts);
            }
        }

        private void AddCsvValue(StringBuilder sb, string value, ref bool hasValue)
        {
            if (hasValue)
                sb.Append(",");

            if (value != null)
            {
                // multilingual character should be quoted, so almost always quotation is necessary
                // if (value.Contains(',') || value.Contains('"') || value.Contains("\n") || value.Contains("\n") )    
                {
                    // We need to quote the value, and therefore get rid of quotes in it
                    // Excel is also fooled by spaces after embedded commas
                    var val = value.Replace("\"", "\"\"");
                    sb.Append("\"");
                    sb.Append(EnforceCsvValueLengthLimit(val));
                    sb.Append("\"");
                }
                // else
                //    sb.Append(EnforceCsvValueLengthLimit(value));
            }

            hasValue = true;
        }

        private static int valueLengthLimit = (int)Math.Pow(2, 15) - 12;
        private string EnforceCsvValueLengthLimit(string value)
        {
            if (value.Length < valueLengthLimit)
                return value;
            else
                return value.Substring(0, valueLengthLimit) + "…";
        }

        #endregion
    }

    public class XstException : Exception
    {
        public XstException(string message) : base(message)
        {
        }
    }
}
