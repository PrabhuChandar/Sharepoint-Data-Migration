using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.IO;
using System.Security;
using System.Windows.Forms;
using spClient = Microsoft.SharePoint.Client;

namespace CopyInfopathFields
{
    public partial class MoveInfopathFields : System.Windows.Forms.Form
    {

        FileStream fsStream;
        static StreamWriter streamWriter = null;
        bool ASMSRequestsListExist = false;
        bool ASMSlibraryExist = false;

        public MoveInfopathFields()
        {
            InitializeComponent();
        }

        #region Event
        private void btnCopy_Click(object sender, EventArgs e)
        {
            string logFileName = string.Empty, strSourceURL = string.Empty, strTargetURL = string.Empty; ;
            try
            {
                logFileName = ConfigurationManager.AppSettings["InfoPathLog"].ToString();

                if (!Directory.Exists(logFileName))
                {
                    Directory.CreateDirectory(logFileName);
                }

                if (txtLogFileName.Text.Contains(".txt"))
                    logFileName += txtLogFileName.Text.Trim();
                else
                    logFileName += txtLogFileName.Text.Trim() + ".txt";

                if (System.IO.File.Exists(logFileName))
                    fsStream = new FileStream(logFileName, FileMode.Append);
                else
                    fsStream = new FileStream(logFileName, FileMode.CreateNew);

                streamWriter = new StreamWriter(fsStream);
                streamWriter.WriteLine("------------------------------------------------------------------------");
                streamWriter.WriteLine("Date and Time Info :" + System.DateTime.Now.ToString("M/d/yyyy hh:mm:ss"));
                streamWriter.WriteLine("------------------------------------------------------------------------");
                strSourceURL = txtSourceURL.Text.Trim();
                strTargetURL = txtDesturl.Text.Trim();
                streamWriter.WriteLine("Source URL :" + strSourceURL);
                streamWriter.WriteLine("Target URL :" + strTargetURL);

                SecureString password = new SecureString();
                foreach (char c in txtPassword.Text.Trim().ToCharArray()) password.AppendChar(c);
                using (spClient.ClientContext SPcontext = new spClient.ClientContext(strTargetURL))
                {
                    SPcontext.Credentials = new SharePointOnlineCredentials(txtUserName.Text.Trim(), password);
                    spClient.Web web = SPcontext.Web;
                    SPcontext.Load(web);
                    spClient.ListCollection listcoll = web.Lists;
                    SPcontext.Load(listcoll);
                    SPcontext.ExecuteQuery();
                    for (int count = 0; count < listcoll.Count; count++)
                    {
                        if (listcoll[count].Title.Equals(txtlistName.Text.Trim()))           //Check whether the newly creating list is already exist
                        {
                            ASMSRequestsListExist = true;
                            break;
                        }
                    }

                    if (!ASMSRequestsListExist)
                    {
                        createListwithfields(SPcontext, txtlistName.Text.Trim());
                        getLibraryData(strSourceURL, strTargetURL);
                    }
                    else
                        getLibraryData(strSourceURL, strTargetURL);
                }

                MessageBox.Show("Completed Successfully!!");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                streamWriter.WriteLine("\r\n\r\n----------------------------- Error -------------------------------");
                streamWriter.WriteLine("Error Message :" + ex.Message);
                streamWriter.WriteLine("Stack Trace :" + ex.StackTrace);
            }

            finally
            {
                streamWriter.Flush();
                streamWriter.Close();
            }
        }
        #endregion

        public void createListwithfields(spClient.ClientContext clientContext, string listName)
        {
            try
            {
                spClient.ListCreationInformation creationInfo = new spClient.ListCreationInformation();
                creationInfo.Title = listName;
                creationInfo.Description = "This list contains InfoPath library data";
                creationInfo.TemplateType = (int)spClient.ListTemplateType.GenericList;
                spClient.List newList = clientContext.Web.Lists.Add(creationInfo);
                clientContext.Load(newList);
                clientContext.ExecuteQuery();
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.ActionPlan + "' Name='" + ASMSRequestConstants.ActionPlan + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.APO + "' Name='" + ASMSRequestConstants.APO + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.APOComments + "' Name='" + ASMSRequestConstants.APOComments + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.AssignedTo + "' Name='" + ASMSRequestConstants.AssignedTo + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='DateTime' DisplayName='" + ASMSRequestConstants.CaseDate + "' Format='DateOnly' Name='" + ASMSRequestConstants.CaseDate + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.CaseNo + "' Name='" + ASMSRequestConstants.CaseNo + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.CaseState + "' Name='" + ASMSRequestConstants.CaseState + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.CaseStatus + "' Name='" + ASMSRequestConstants.CaseStatus + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.CaseURL + "' Name='" + ASMSRequestConstants.CaseURL + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.Cost + "' Name='" + ASMSRequestConstants.Cost + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.DisapprovedReason + "' Name='" + ASMSRequestConstants.DisapprovedReason + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='DateTime' DisplayName='" + ASMSRequestConstants.DOC + "' Format='DateOnly' Name='" + ASMSRequestConstants.DOC + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='DateTime' DisplayName='" + ASMSRequestConstants.ImplementedDate + "' Format='DateOnly' Name='" + ASMSRequestConstants.ImplementedDate + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.IsApproved + "' Name='" + ASMSRequestConstants.IsApproved + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.IsConfident + "' Name='" + ASMSRequestConstants.IsConfident + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.IsNewCase + "' Name='" + ASMSRequestConstants.IsNewCase + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.IsVerified + "' Name='" + ASMSRequestConstants.IsVerified + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.mComments + "' Name='" + ASMSRequestConstants.mComments + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.NinetyDay + "' Name='" + ASMSRequestConstants.NinetyDay + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.Notes + "' Name='" + ASMSRequestConstants.Notes + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.NotRechecked + "' Name='" + ASMSRequestConstants.NotRechecked + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.ParentCase + "' Name='" + ASMSRequestConstants.ParentCase + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.PrjtNo + "' Name='" + ASMSRequestConstants.PrjtNo + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.Recheck + "' Name='" + ASMSRequestConstants.Recheck + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.RecheckComments + "' Name='" + ASMSRequestConstants.RecheckComments + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.RecheckDate + "' Name='" + ASMSRequestConstants.RecheckDate + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.ReportedBy + "' Name='" + ASMSRequestConstants.ReportedBy + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.ReqDesc + "' Name='" + ASMSRequestConstants.ReqDesc + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.RequestedBy + "' Name='" + ASMSRequestConstants.RequestedBy + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.Requester + "' Name='" + ASMSRequestConstants.Requester + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='DateTime' DisplayName='" + ASMSRequestConstants.ResolutionDueDate + "' Format='DateOnly' Name='" + ASMSRequestConstants.ResolutionDueDate + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.StatusDate + "' Name='" + ASMSRequestConstants.StatusDate + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.Summary + "' Name='" + ASMSRequestConstants.Summary + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.Urgency + "' Name='" + ASMSRequestConstants.Urgency + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.VerifiedDate + "' Name='" + ASMSRequestConstants.VerifiedDate + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.VerifiedReason + "' Name='" + ASMSRequestConstants.VerifiedReason + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.WFState + "' Name='" + ASMSRequestConstants.WFState + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.CreatedBy + "' Name='" + ASMSRequestConstants.CreatedBy + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.ModifiedBy + "' Name='" + ASMSRequestConstants.ModifiedBy + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);
                newList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='" + ASMSRequestConstants.AMSRequestWorkflow + "' Name='" + ASMSRequestConstants.AMSRequestWorkflow + "' />", true, spClient.AddFieldOptions.AddToDefaultContentType);


                clientContext.Load(newList);
                clientContext.ExecuteQuery();
                newList.Update();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                streamWriter.WriteLine("\r\n\r\n----------------------------- Error -------------------------------");
                streamWriter.WriteLine("Error Message :" + ex.Message);
                streamWriter.WriteLine("Stack Trace :" + ex.StackTrace);
            }
        }

        public void AddItemsIntoList(spClient.ListItem Item, spClient.ClientContext clientContext, spClient.ListItemCollection collListItem, spClient.ListItem oListItem, List splist)
        {
            string strAuthorName = string.Empty, strEditorName = string.Empty;
            try
            {
                oListItem[ListInternalName.Title] = Convert.ToString(Item[InternalFieldName.Name]).Replace(".xml", "");
                oListItem[ListInternalName.ActionPlan] = Convert.ToString(Item[InternalFieldName.ActionPlan]);
                oListItem[ListInternalName.APO] = Convert.ToString(Item[InternalFieldName.APO]);
                oListItem[ListInternalName.APOComments] = Convert.ToString(Item[InternalFieldName.APOComments]);
                oListItem[ListInternalName.AssignedTo] = Convert.ToString(Item[InternalFieldName.AssignedTo]);

                if (string.IsNullOrEmpty(Convert.ToString(Item[InternalFieldName.CaseDate])))
                    oListItem[ListInternalName.CaseDate] = null;
                else
                    oListItem[ListInternalName.CaseDate] = Item[InternalFieldName.CaseDate];

                oListItem[ListInternalName.CaseNo] = Convert.ToString(Item[InternalFieldName.CaseNo]);
                oListItem[ListInternalName.CaseState] = Convert.ToString(Item[InternalFieldName.CaseState]);
                oListItem[ListInternalName.CaseStatus] = Convert.ToString(Item[InternalFieldName.CaseStatus]);
                oListItem[ListInternalName.CaseURL] = Convert.ToString(Item[InternalFieldName.CaseURL]);
                oListItem[ListInternalName.Cost] = Convert.ToString(Item[InternalFieldName.Cost]);
                oListItem[ListInternalName.DisapprovedReason] = Convert.ToString(Item[InternalFieldName.DisapprovedReason]);

                if (string.IsNullOrEmpty(Convert.ToString(Item[InternalFieldName.DOC])))
                    oListItem[ListInternalName.DOC] = null;
                else
                    oListItem[ListInternalName.DOC] = Item[InternalFieldName.DOC];

                if (string.IsNullOrEmpty(Convert.ToString(Item[InternalFieldName.ImplementedDate])))
                    oListItem[ListInternalName.ImplementedDate] = null;
                oListItem[ListInternalName.ImplementedDate] = Item[InternalFieldName.ImplementedDate];

                oListItem[ListInternalName.IsApproved] = Convert.ToString(Item[InternalFieldName.IsApproved]);
                oListItem[ListInternalName.IsConfident] = Convert.ToString(Item[InternalFieldName.IsConfident]);
                oListItem[ListInternalName.IsNewCase] = Convert.ToString(Item[InternalFieldName.IsNewCase]);
                oListItem[ListInternalName.IsVerified] = Convert.ToString(Item[InternalFieldName.IsVerified]);
                oListItem[ListInternalName.mComments] = Convert.ToString(Item[InternalFieldName.mComments]);
                oListItem[ListInternalName.Notes] = Convert.ToString(Item[InternalFieldName.Notes]);

                oListItem[ListInternalName.ParentCase] = Convert.ToString(Item[InternalFieldName.ParentCase]);
                oListItem[ListInternalName.PrjtNo] = Convert.ToString(Item[InternalFieldName.PrjtNo]);
                oListItem[ListInternalName.Recheck] = Convert.ToString(Item[InternalFieldName.Recheck]);
                oListItem[ListInternalName.RecheckComments] = Convert.ToString(Item[InternalFieldName.RecheckComments]);
                oListItem[ListInternalName.RecheckDate] = Convert.ToString(Item[InternalFieldName.RecheckDate]);
                oListItem[ListInternalName.ReportedBy] = Convert.ToString(Item[InternalFieldName.ReportedBy]);
                oListItem[ListInternalName.ReqDesc] = Convert.ToString(Item[InternalFieldName.ReqDesc]);
                oListItem[ListInternalName.RequestedBy] = Convert.ToString(Item[InternalFieldName.RequestedBy]);
                oListItem[ListInternalName.Requester] = Convert.ToString(Item[InternalFieldName.Requester]);

                if (string.IsNullOrEmpty(Convert.ToString(Item[InternalFieldName.ResolutionDueDate])))
                    oListItem[ListInternalName.ResolutionDueDate] = null;
                else
                    oListItem[ListInternalName.ResolutionDueDate] = Item[InternalFieldName.ResolutionDueDate];

                oListItem[ListInternalName.StatusDate] = Convert.ToString(Item[InternalFieldName.StatusDate]);
                oListItem[ListInternalName.Summary] = Convert.ToString(Item[InternalFieldName.Summary]);
                oListItem[ListInternalName.Urgency] = Convert.ToString(Item[InternalFieldName.Urgency]);
                oListItem[ListInternalName.VerifiedDate] = Convert.ToString(Item[InternalFieldName.VerifiedDate]);
                oListItem[ListInternalName.VerifiedReason] = Convert.ToString(Item[InternalFieldName.VerifiedReason]);

                #region //Seperate content Type and it is being a single line text box
                oListItem[ListInternalName.NinetyDay] = Convert.ToString(Item[InternalFieldName.NinetyDay]);
                oListItem[ListInternalName.NotRechecked] = Convert.ToString(Item[InternalFieldName.NotRechecked]);
                oListItem[ListInternalName.WFState] = Convert.ToString(Item[InternalFieldName.WFState]);

                switch (Convert.ToString(Item[InternalFieldName.AMSReque]))
                {
                    case "":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "";
                        break;
                    case "0":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Not Started";
                        break;
                    case "1":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Failed On Start";
                        break;
                    case "2":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "In Progress";
                        break;
                    case "3":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Error Occurred";
                        break;
                    case "4":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Stopped By User";
                        break;
                    case "5":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Completed";
                        break;
                    case "6":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Failed On Start Retrying ";
                        break;
                    case "7":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Error Occurred Retrying ";
                        break;
                    case "8":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "ViewQuery Overflow";
                        break;
                    case "15":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Canceled";
                        break;
                    case "16":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Approved";
                        break;
                    case "17":
                        oListItem[ListInternalName.AMSRequestWorkflow] = "Rejected";
                        break;
                }

                #endregion
                FieldUserValue oValue = Item["Author"] as FieldUserValue;                                    //get Created by from the Infopath library
                strAuthorName = oValue.LookupValue;
                oListItem[ListInternalName.CreatedBy] = Convert.ToString(strAuthorName);
                FieldUserValue Modifiedby = Item["Editor"] as FieldUserValue;                                //get Modified by from the Infopath library
                strEditorName = Modifiedby.LookupValue;
                oListItem[ListInternalName.ModifiedBy] = Convert.ToString(strEditorName);
                oListItem.Update();
                clientContext.Load(oListItem);
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message + "on the method AddItemsIntoList");
                streamWriter.WriteLine("\r\n\r\n----------------------------- Error -------------------------------");
                streamWriter.WriteLine("Error Message :" + ex.Message);
                streamWriter.WriteLine("Stack Trace :" + ex.StackTrace);
            }
        }

        public void getLibraryData(string strSourceURL, string strTargetURL)
        {
            try
            {
                using (spClient.ClientContext sourceContext = new spClient.ClientContext(strSourceURL)) //Site URL of InfoPath Libray
                {
                    Web web = sourceContext.Web;
                    spClient.ListCollection listColl = web.Lists;
                    sourceContext.Load(listColl);
                    sourceContext.ExecuteQuery();
                    for (int Count = 0; Count < listColl.Count; Count++)
                    {
                        if (listColl[Count].Title.Equals(txtLibraryName.Text.Trim()))
                        {
                            ASMSlibraryExist = true;
                            break;
                        }
                    }
                    if (ASMSlibraryExist)
                    {
                        spClient.List oList = sourceContext.Web.Lists.GetByTitle(txtLibraryName.Text.Trim());
                        spClient.CamlQuery camlQuery = new spClient.CamlQuery();
                        camlQuery.ViewXml = "<Query><OrderBy><FieldRef Name='ID' /></OrderBy></Query>";
                        spClient.ListItemCollection collListItem = oList.GetItems(camlQuery);

                        sourceContext.Load(collListItem);
                        sourceContext.ExecuteQuery();

                        SecureString password = new SecureString();
                        foreach (char c in txtPassword.Text.Trim().ToCharArray()) password.AppendChar(c);
                        using (spClient.ClientContext descontext = new spClient.ClientContext(txtDesturl.Text.Trim()))
                        {
                            descontext.Credentials = new SharePointOnlineCredentials(txtUserName.Text.Trim(), password);
                            spClient.Web oweb = descontext.Web;
                            spClient.ListCollection listcoll = oweb.Lists;

                            spClient.List spList = descontext.Web.Lists.GetByTitle(txtlistName.Text);
                            spClient.ListItemCreationInformation itemCreateInfo = new spClient.ListItemCreationInformation();
                            descontext.Load(spList);
                            descontext.ExecuteQuery();

                            for (int i = 0; i < collListItem.Count; i++)
                            {
                                try
                                {
                                    spClient.ListItem oListItem = spList.AddItem(itemCreateInfo);
                                    sourceContext.Load(collListItem[i]);
                                    AddItemsIntoList(collListItem[i], descontext, collListItem, oListItem, spList);
                                }
                                catch (Exception ex)
                                {
                                    //   MessageBox.Show(ex.Message);
                                    streamWriter.WriteLine("\r\n\r\n----------------------------- Error -------------------------------");
                                    streamWriter.WriteLine("Error Message :" + ex.Message);
                                    streamWriter.WriteLine("Stack Trace :" + ex.StackTrace);
                                }
                            }
                            sourceContext.ExecuteQuery();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Library does not exist. Please give valid InfoPath Library Name. ");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                streamWriter.WriteLine("\r\n\r\n----------------------------- Error -------------------------------");
                streamWriter.WriteLine("Error Message :" + ex.Message);
                streamWriter.WriteLine("Stack Trace :" + ex.StackTrace);
            }
        }

        private void txtlistName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtLibraryName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtDesturl_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void txtSourceURL_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void txtLogFileName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtUserName_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
