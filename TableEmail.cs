using MSOutlook = Microsoft.Office.Interop.Outlook;
using MySqlConnector;


namespace OutlookToMariadb
{
	public class TableEmail
	{
		public MySqlServer DB;

		public MySqlParameter par_folder = new MySqlParameter("@folder", MySqlDbType.VarChar, 200);
		public MySqlParameter par_storeid = new MySqlParameter("@storeid", MySqlDbType.VarChar, 255);
		public MySqlParameter par_bcc = new MySqlParameter("@bcc", MySqlDbType.VarChar, 500);
		public MySqlParameter par_attachments = new MySqlParameter("@attachments", MySqlDbType.Text);
		public MySqlParameter par_body = new MySqlParameter("@body", MySqlDbType.MediumText);
		public MySqlParameter par_bodyformat = new MySqlParameter("@bodyformat", MySqlDbType.VarChar, 50);
		public MySqlParameter par_cc = new MySqlParameter("@cc", MySqlDbType.VarChar, 500);
		public MySqlParameter par_creationtime = new MySqlParameter("@creationtime", MySqlDbType.DateTime);
		public MySqlParameter par_deferreddeliverytime = new MySqlParameter("@deferreddeliverytime", MySqlDbType.DateTime);
		public MySqlParameter par_entryid = new MySqlParameter("@entryid", MySqlDbType.VarChar, 255);
		public MySqlParameter par_htmlbody = new MySqlParameter("@htmlbody", MySqlDbType.MediumText);
		public MySqlParameter par_importance = new MySqlParameter("@importance", MySqlDbType.VarChar, 20);
		public MySqlParameter par_internetcodepage = new MySqlParameter("@internetcodepage", MySqlDbType.Int64);
		public MySqlParameter par_lastmodificationtime = new MySqlParameter("@lastmodificationtime", MySqlDbType.DateTime);
		public MySqlParameter par_messageclass = new MySqlParameter("@messageclass", MySqlDbType.VarChar, 255);
		public MySqlParameter par_readreceiptrequested = new MySqlParameter("@readreceiptrequested", MySqlDbType.Bit);
		public MySqlParameter par_receivedbyentryid = new MySqlParameter("@receivedbyentryid", MySqlDbType.VarChar, 255);
		public MySqlParameter par_receivedbyname = new MySqlParameter("@receivedbyname", MySqlDbType.VarChar, 255);
		public MySqlParameter par_receivedonbehalfofentryid = new MySqlParameter("@receivedonbehalfofentryid", MySqlDbType.VarChar, 255);
		public MySqlParameter par_receivedtime = new MySqlParameter("@receivedtime", MySqlDbType.DateTime);
		public MySqlParameter par_recipients = new MySqlParameter("@recipients", MySqlDbType.MediumText);
		public MySqlParameter par_replyrecipients = new MySqlParameter("@replyrecipients", MySqlDbType.MediumText);
		public MySqlParameter par_rtfbody = new MySqlParameter("@rtfbody", MySqlDbType.MediumText);
		public MySqlParameter par_senderemailaddress = new MySqlParameter("@senderemailaddress", MySqlDbType.VarChar, 255);
		public MySqlParameter par_senderemailtype = new MySqlParameter("@senderemailtype", MySqlDbType.VarChar, 255);
		public MySqlParameter par_sendername = new MySqlParameter("@sendername", MySqlDbType.VarChar, 255);
		public MySqlParameter par_sentonbehalfOfName = new MySqlParameter("@sentonbehalfOfName", MySqlDbType.MediumText);
		public MySqlParameter par_size = new MySqlParameter("@size", MySqlDbType.Int64);
		public MySqlParameter par_subject = new MySqlParameter("@subject", MySqlDbType.MediumText);
		public MySqlParameter par_to = new MySqlParameter("@to", MySqlDbType.VarChar, 2000);
		public MySqlParameter par_unread = new MySqlParameter("@unread", MySqlDbType.Bit);

		public TableEmail(MySqlServer db)
		{
			DB = db;
		}

		public List<Tuple<string, string>> GetKeyOfEmails()
        {
			var res = new List<Tuple<string, string>>();

			var conn = DB.Connect();
			var cmd = conn.CreateCommand();
			cmd.CommandText = $"select storeid, entryid from email";
			var reader = cmd.ExecuteReader();

			while(reader.Read())
				res.Add(new Tuple<string, string>(reader.GetString(0), reader.GetString(1)));

			reader.Close();
			conn.Close();
			return res;
		}

		public List<Tuple<string, string, string>> GetActions()
		{
			var res = new List<Tuple<string, string, string>>();

			var conn = DB.Connect();
			var cmd = conn.CreateCommand();
			cmd.CommandText = $"select storeid, entryid, action from email where action is not null";
			var reader = cmd.ExecuteReader();

			while (reader.Read())
				res.Add(new Tuple<string, string, string>(reader.GetString(0), reader.GetString(1), reader.GetString(2)));

			reader.Close();
			conn.Close();
			return res;
		}

		public int DeleteByStoreIdAndEntryIds(List<Tuple<string, string>> ids)
		{
			int res = -1;

			var conn = DB.Connect();
			var cmd = conn.CreateCommand();
			foreach (var id in ids) { 
				cmd.CommandText = $"delete from email where storeid='{id.Item1}' and entryid='{id.Item2}';";
				cmd.ExecuteNonQuery();
			}
			
			conn.Close();
			return res;
        }

		public int BatchInsertEmails(List<Mail> mails)
        {
			int res = 0;
			int errors = 0;
			var conn = DB.Connect();
			var cmd = CreateInsertCommand(conn);

			foreach(var m in mails)
            {
				var v = Insert(m.Folder, m.MailItem, cmd);
				if (v)
					res++;
				else
					Console.WriteLine("Insert to table email error");

				var s = $"\r{res} of {mails.Count}, errors：{errors}";
				Console.Write(s);
			}


			return res;
        }

		public MySqlCommand CreateInsertCommand(MySqlConnection conn, MySqlTransaction? trans = null)
		{
			MySqlCommand cmd = trans == null ? new MySqlCommand(sql_insert, conn) : new MySqlCommand(sql_insert, conn, trans);

			cmd.Parameters.Add(par_folder);
			cmd.Parameters.Add(par_storeid);
			cmd.Parameters.Add(par_bcc);
			cmd.Parameters.Add(par_attachments);
			cmd.Parameters.Add(par_body);
			cmd.Parameters.Add(par_bodyformat);
			cmd.Parameters.Add(par_cc);
			cmd.Parameters.Add(par_creationtime);
			cmd.Parameters.Add(par_deferreddeliverytime);
			cmd.Parameters.Add(par_entryid);
			cmd.Parameters.Add(par_htmlbody);
			cmd.Parameters.Add(par_importance);
			cmd.Parameters.Add(par_internetcodepage);
			cmd.Parameters.Add(par_lastmodificationtime);
			cmd.Parameters.Add(par_messageclass);
			cmd.Parameters.Add(par_readreceiptrequested);
			cmd.Parameters.Add(par_receivedbyentryid);
			cmd.Parameters.Add(par_receivedbyname);
			cmd.Parameters.Add(par_receivedonbehalfofentryid);
			cmd.Parameters.Add(par_receivedtime);
			cmd.Parameters.Add(par_recipients);
			cmd.Parameters.Add(par_replyrecipients);
			cmd.Parameters.Add(par_rtfbody);
			cmd.Parameters.Add(par_senderemailaddress);
			cmd.Parameters.Add(par_senderemailtype);
			cmd.Parameters.Add(par_sendername);
			cmd.Parameters.Add(par_sentonbehalfOfName);
			cmd.Parameters.Add(par_size);
			cmd.Parameters.Add(par_subject);
			cmd.Parameters.Add(par_to);
			cmd.Parameters.Add(par_unread);

			return cmd;
		}

		public bool Insert(MSOutlook.MAPIFolder folder, MSOutlook.MailItem item, MySqlCommand insertCommand)
		{
			par_folder.Value = folder.FullFolderPath;
			par_storeid.Value = folder.StoreID;

			par_bcc.Value = item.BCC;
			//par_attachments.Value = item.Attachments.ToString();
			//par_body.Value = ToBase64String(item.Body);
			par_bodyformat.Value = item.BodyFormat.ToString();
			par_cc.Value = item.CC;
			par_creationtime.Value = item.CreationTime;
			par_deferreddeliverytime.Value = item.DeferredDeliveryTime;
			par_entryid.Value = item.EntryID;
			//par_htmlbody.Value = ToBase64String(item.HTMLBody);
			par_importance.Value = item.Importance.ToString();
			par_internetcodepage.Value = item.InternetCodepage;
			par_lastmodificationtime.Value = item.LastModificationTime;
			par_messageclass.Value = item.MessageClass;
			par_readreceiptrequested.Value = item.ReadReceiptRequested;
			par_receivedbyentryid.Value = item.ReceivedByEntryID;
			par_receivedbyname.Value = item.ReceivedByName;
			par_receivedonbehalfofentryid.Value = item.ReceivedOnBehalfOfEntryID;
			par_receivedtime.Value = item.ReceivedTime;
			//par_recipients.Value = item.Recipients.ToString();
			//par_replyrecipients.Value = item.ReplyRecipients.ToString();
			//par_rtfbody.Value = item.RTFBody;
			par_senderemailaddress.Value = item.SenderEmailAddress;
			par_senderemailtype.Value = item.SenderEmailType;
			par_sendername.Value = item.SenderName;
			par_sentonbehalfOfName.Value = item.SentOnBehalfOfName;
			par_size.Value = item.Size;
			par_subject.Value = Utils.ToBase64String(item.Subject);
			par_to.Value = item.To;
			par_unread.Value = item.UnRead;

			// handle attachments
			par_attachments.Value = null;
			if (item.Attachments != null && item.Attachments.Count != 0)
			{
				string an = "[";
				for (int a = 1; a < item.Attachments.Count; a++)
				{
					var attach = item.Attachments[a];
					if (a != 1)
						an += ",";
					an += ($"{{\"文件名\": \"{attach.FileName}\",  \"显示为\": \"{attach.DisplayName}\", \"大小\": \"{attach.Size}\"}}");
				}
				an += "]";
				par_attachments.Value = an;
			}

			//handle recipients;
			par_recipients.Value = null;
			if (item.Recipients != null && item.Recipients.Count != 0)
			{
				string an = "";
				for (int i = 1; i < item.Recipients.Count; i++)
				{
					an += item.Recipients[i].Name + ";";
				}
				par_recipients.Value = an;
			}

			//handle replyrecipients;
			par_replyrecipients.Value = null;
			if (item.ReplyRecipients != null && item.ReplyRecipients.Count != 0)
			{
				string an = "";
				for (int i = 1; i < item.ReplyRecipients.Count; i++)
				{
					an += item.ReplyRecipients[i].Name + ";";
				}
				par_replyrecipients.Value = an;
			}


			// handle mail body
			par_rtfbody.Value = null;
			par_htmlbody.Value = null;
			par_body.Value = null;

			switch (item.BodyFormat)
			{
				case MSOutlook.OlBodyFormat.olFormatPlain:
					par_body.Value = item.Body;
					break;

				case MSOutlook.OlBodyFormat.olFormatHTML:
					par_htmlbody.Value = Utils.ToBase64String(item.HTMLBody);
					break;

				case MSOutlook.OlBodyFormat.olFormatRichText:
					par_rtfbody.Value = item.RTFBody;
					break;

				case MSOutlook.OlBodyFormat.olFormatUnspecified:
					par_body.Value = Utils.ToBase64String(item.Body);
					break;
			}

			insertCommand.ExecuteNonQuery();

			return true;
		}

		public static string sql_insert = @"Insert into email (
                                            `folder`, 
											`storeid`,
                                            `bcc`, 
                                            `attachments`, 
                                            `body`, 
                                            `bodyformat`,
                                            `cc`,
	                                        `creationtime`,
	                                        `deferreddeliverytime`,
	                                        `entryid`,
	                                        `htmlbody`,
	                                        `importance`,
	                                        `internetcodepage`,
	                                        `lastmodificationtime`,
	                                        `messageclass` ,
	                                        `readreceiptrequested`,
	                                        `receivedbyentryid` ,
	                                        `receivedbyname` ,
	                                        `receivedonbehalfofentryid` ,
	                                        `receivedtime` ,
	                                        `recipients` ,
	                                        `replyrecipients`,
	                                        `rtfbody` ,
	                                        `senderemailaddress` ,
	                                        `senderemailtype` ,
	                                        `sendername` ,
	                                        `sentonbehalfOfName` ,
	                                        `size` ,
	                                        `subject`,
	                                        `to` ,
	                                        `unread`) values(
                                            @folder, 
                                            @storeid, 
                                            @bcc, 
                                            @attachments, 
                                            @body, 
                                            @bodyformat,
                                            @cc,
	                                        @creationtime,
	                                        @deferreddeliverytime,
	                                        @entryid,
	                                        @htmlbody,
	                                        @importance,
	                                        @internetcodepage,
	                                        @lastmodificationtime,
	                                        @messageclass ,
	                                        @readreceiptrequested,
	                                        @receivedbyentryid ,
	                                        @receivedbyname ,
	                                        @receivedonbehalfofentryid ,
	                                        @receivedtime ,
	                                        @recipients ,
	                                        @replyrecipients,
	                                        @rtfbody ,
	                                        @senderemailaddress ,
	                                        @senderemailtype ,
	                                        @sendername ,
	                                        @sentonbehalfOfName ,
	                                        @size ,
	                                        @subject,
	                                        @to ,
	                                        @unread)";

	}
}
