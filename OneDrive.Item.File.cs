//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using Microsoft.Graph;
using System.Text.RegularExpressions;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace Cliver
{
    public partial class OneDrive
    {
        public class File : Item
        {
            public static File New(OneDrive oneDrive, string remoteFile, bool createIfNotExists)
            {
                Item item = oneDrive.GetItemByPath(remoteFile);
                if (item != null)
                {
                    if (item is File)
                        return (File)item;
                    throw new Exception("Remote path points to not a file: " + remoteFile);
                }
                if (!createIfNotExists)
                    return null;

                Match m = Regex.Match(remoteFile, @"(?'ParentFolder'.*)[\\\/]+(?'Name'.*)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (!m.Success)
                    throw new Exception("Remote file path could not be separated: " + remoteFile);

                Folder parentFolder = Folder.New(oneDrive, m.Groups["ParentFolder"].Value, true);
                DriveItem di = new DriveItem
                {
                    Name = m.Groups["Name"].Value,
                    File = new Microsoft.Graph.File
                    {
                    },
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"@microsoft.graph.conflictBehavior", "rename"}
                    }
                };
                DriveItem driveItem = Task.Run(() =>
                {
                    return parentFolder.DriveItemRequestBuilder.Children.Request().AddAsync(di);
                }).Result;
                return new File(oneDrive, driveItem);
            }

            internal File(OneDrive oneDrive, DriveItem driveItem) : base(oneDrive, driveItem)
            {
            }

            //public bool IsCheckInSupported
            //{
            //    get
            //    {
            //        if (checkInIsSupported == null)
            //            checkInIsSupported = GetCheckStatus() != CheckStatus.NotSupported;
            //        return (bool)checkInIsSupported;
            //    }
            //}

            public enum CheckStatus
            {
                NotSupported,
                CheckedOutByNotMe,
                CheckedIn,
                CheckedOut,
            }
            public CheckStatus GetCheckStatus()
            {
                if (checkInIsSupported == false)
                    return CheckStatus.NotSupported;

                var i = Task.Run(() =>
                {
                    return DriveItemRequestBuilder.Request().Select("id, publication").GetAsync();
                }).Result;
                //Log.Debug0(i.ToStringByJson());
                if (i.Publication == null)//if NULL then checkout is not supported
                {
                    checkInIsSupported = false;
                    return CheckStatus.NotSupported;
                }
                checkInIsSupported = true;
                string s = i.Publication.Level.ToLower();
                if (s == "published")
                    return CheckStatus.CheckedIn;
                if (s == "checkout")
                    return CheckStatus.CheckedOut;
                throw new Exception("Unknown Publication.Level: " + s);
            }
            bool? checkInIsSupported = null;

            /// <summary>
            /// (!)Not supported on a personal OneDrive: https://learn.microsoft.com/en-us/answers/questions/574546/is-checkin-checkout-files-supported-by-onedrive-pe.html
            /// </summary>
            /// <param name="throwExceptionIfFailed"></param>
            public CheckStatus CheckOut(bool throwExceptionIfFailed = false)
            {
                CheckStatus cs = GetCheckStatus();
                if (cs == CheckStatus.NotSupported)
                    return cs;
                if (cs == CheckStatus.CheckedOut && CheckIn() != CheckStatus.CheckedIn)
                    if (throwExceptionIfFailed)
                        throw new Exception(Cliver.Log.GetThisMethodName() + " failed on the file:\r\n" + DriveItem.WebUrl + "\r\nCheck status of the file: " + CheckStatus.CheckedOutByNotMe.ToString());
                    else
                        return CheckStatus.CheckedOutByNotMe;
                //if (cs == CheckStatus.CheckedOut)//must work on sharepoint
                //{//get who the item is checked out by:
                //    if (SharepointIds == null)
                //        throw new Exception("SharepointIds are NULL while the DriveItem status is CheckedOut.");
                //    //check if the item is checkedout by someone else
                //    FieldValueSet fieldValueSet = Task.Run(() =>
                //    {
                //        return OneDrive.Client.Sites[SharepointIds.SiteId].Lists[ListItem.Id].Items[ItemId].Fields.Request().Expand("CheckoutUser").GetAsync();
                //    }).Result;
                //    Log.Debug0(fieldValueSet.AdditionalData.ToStringByJson());
                //    object checkoutUser = fieldValueSet.AdditionalData["CheckoutUser"];
                //    if (checkoutUser == null)
                //        throw new Exception("Could not get checkoutUser for the DriveItem.");
                //    //who checked out???
                //    //if(checkoutUser is Me)
                //    //    return CheckStatus.CheckedOut;
                //    if (throwExceptionIfFailed)
                //        throw new Exception(Cliver.Log.GetThisMethodName() + " failed on the file:\r\n" + DriveItem.WebUrl
                //            + "\r\nThe file is checked out by user: " + checkoutUser.ToString()
                //            );
                //    else
                //        return CheckStatus.CheckedOutByNotMe;
                //}

                Task.Run(() =>
                {
                    DriveItemRequestBuilder.Checkout().Request().PostAsync();//not supported for a personal OneDrive: https://learn.microsoft.com/en-us/answers/questions/574546/is-checkin-checkout-files-supported-by-onedrive-pe.html
                }).Wait();

                SleepRoutines.WaitForCondition(() =>
                {
                    cs = GetCheckStatus();
                    return cs == CheckStatus.CheckedIn;
                }, CheckStatusChangeTimeoutSecs * 1000, 1000);
                if (cs != CheckStatus.CheckedOut && throwExceptionIfFailed)
                    throw new Exception(Cliver.Log.GetThisMethodName() + " failed on the file:\r\n" + DriveItem.WebUrl + "\r\nCheck status of the file: " + cs.ToString());

                return cs;
            }

            List<string> GetCurrentEditors()
            {
                //get who keeps it open (for Excel sheets):                    
                DriveItem di = GetDriveItem(null, "activities");
                Log.Debug0(di.AdditionalData.ToStringByJson());

                Log.Debug0(SharepointIds.ToStringByJson());
                //Log.Debug0(ListItem.SharepointIds.ToStringByJson());

                object activities = di.AdditionalData["activities"];
                Log.Debug0(activities.GetType().ToString());
                Log.Debug0(ListItem.AdditionalData.ToStringByJson());

                var t = Task.Run(() =>
                {
                    //!!!GetActivitiesByInterval gives not user names
                    return OneDrive.Client.Me.Drives[DriveId].Items[ItemId].GetActivitiesByInterval(DateTime.Now.AddMinutes(-20).ToString("yyyy-MM-dd hh:mm:ss"), DateTime.Now.AddMinutes(2).ToString("yyyy-MM-dd hh:mm:ss"), "hour").Request().GetAsync();
                }).Result;
                Log.Debug0(t.ToStringByJson());

                FieldValueSet fieldValueSet = Task.Run(() =>
                {
                    var queryOptions = new List<QueryOption>() { new QueryOption("expand", "activities") };
                    return OneDrive.Client.Sites[SharepointIds.SiteId].Lists[SharepointIds.ListItemId].Items[ItemId].Fields.Request(queryOptions).GetAsync();
                }).Result;//!!!The problem seems to be because of missing oAuth permissions for Sites on the client.
                Log.Debug0(fieldValueSet.AdditionalData.ToStringByJson());

                return new List<string> { "test" };
            }

            /// <summary>
            /// Default time to wait for the check status value to change after check-in and check-out. 
            /// Sometimes it seems to need time to change.
            /// </summary>
            public int CheckStatusChangeTimeoutSecs = 0;

            /// <summary>
            /// (!)Not supported on a personal OneDrive: https://learn.microsoft.com/en-us/answers/questions/574546/is-checkin-checkout-files-supported-by-onedrive-pe.html
            /// </summary>
            /// <param name="comment"></param>
            /// <param name="throwExceptionIfFailed"></param>
            public CheckStatus CheckIn(string comment = null, bool throwExceptionIfFailed = false)
            {
                if (GetCheckStatus() == CheckStatus.NotSupported)
                    return CheckStatus.NotSupported;

                if (comment == null)
                    comment = "by " + Log.ProgramName;
                Task.Run(() =>
                {
                    DriveItemRequestBuilder.Checkin(/*"published"*/null, comment).Request().PostAsync();//not supported for a personal OneDrive: https://learn.microsoft.com/en-us/answers/questions/574546/is-checkin-checkout-files-supported-by-onedrive-pe.html
                }).Wait();

                CheckStatus cs = CheckStatus.NotSupported;
                SleepRoutines.WaitForCondition(() =>
                {
                    cs = GetCheckStatus();
                    return cs == CheckStatus.CheckedIn;
                }, CheckStatusChangeTimeoutSecs * 1000, 1000);
                if (cs != CheckStatus.CheckedIn && throwExceptionIfFailed)
                    throw new Exception(Cliver.Log.GetThisMethodName() + " failed on the file:\r\n" + DriveItem.WebUrl + "\r\nCheck status of the file: " + cs.ToString());
                return cs;
            }

            public string Download2Folder(string localFolder, string localFileName = null)
            {
                if (localFileName == null)
                    localFileName = DriveItem.Name;
                string localFile = localFolder + Path.DirectorySeparatorChar + localFileName;
                Download(localFile);
                return localFile;
            }

            public void Download(string localFile)
            {
                using (Stream s = Task.Run(() =>
                {
                    return DriveItemRequestBuilder.Content.Request().GetAsync();
                }).Result
                    )
                {
                    using (var fileStream = System.IO.File.Create(localFile))
                    {
                        s.Seek(0, SeekOrigin.Begin);
                        s.CopyTo(fileStream);
                    }
                }
            }

            public void Upload(string localFile)
            {
                using (Stream s = System.IO.File.OpenRead(localFile))
                {
                    DriveItem = Task.Run(() =>
                    {
                        return DriveItemRequestBuilder.Content.Request().PutAsync<DriveItem>(s);
                    }).Result;
                }
            }

            public Folder GetFolder()
            {
                return (Folder)GetParent();
            }
        }
    }
}