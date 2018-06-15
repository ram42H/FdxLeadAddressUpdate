using FdxLeadAssignmentPlugin;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json; 
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FdxLeadAddressUpdate
{
    public class LeadAddressUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            string DEV_ENVIRONMENT_URL = "http://SMARTCRMSync.1800dentist.com/api";
            string STAGE_ENVIRONMENT_URL = "http://SMARTCRMSyncStage.1800dentist.com/api";
            string PROD_ENVIRONMENT_URL = "http://SMARTCRMSyncProd.1800dentist.com/api";
            string smartCrmSyncWebServiceUrl = STAGE_ENVIRONMENT_URL;

            //Extract the tracing service for use in debugging sandboxed plug-ins....
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution contest from the service provider....
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            int step = 0;
            string apiParmCreate = "";
            string apiParmUpdate = "";
            string apiParmAccCreate = "";
            string url = "";
            bool updateAccount_GMNo = false;
            bool createNewGMNo = false;

            //Call Input parameter collection to get all the data passes....
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.Depth == 1)
            {
                Entity leadEntity = (Entity)context.InputParameters["Target"];

                if (leadEntity.LogicalName != "lead")
                    return;

                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                    IOrganizationService impersonatedService = serviceFactory.CreateOrganizationService(null);

                    step = 0;
                    Guid accountid = Guid.Empty;

                    tracingService.Trace("fdx_accountcontext: {0}", leadEntity.Contains("fdx_accountcontext"));
                    //prevents account address update plugin to run below code
                    if (!leadEntity.Attributes.Contains("fdx_accountcontext"))
                    {
                        if (leadEntity.Attributes.Contains("parentaccountid"))
                        {
                            step = 1;
                            //Case 1: Lead when tagged with account
                            if (leadEntity.Attributes["parentaccountid"] != null)
                            {
                                step = 2;
                                #region Lead when tagged with account...
                                ColumnSet columns = new ColumnSet("name", "fdx_goldmineaccountnumber", "fdx_gonogo", "address1_line1", "address1_line2", "address1_city",
                                        "fdx_stateprovinceid", "fdx_zippostalcodeid", "telephone1", "address1_country",
                                        "fdx_prospectgroup", "defaultpricelevelid", "fdx_prospectpriority", "fdx_prospectscore", "fdx_prospectpercentile", "fdx_ratesource", "fdx_pprrate", "fdx_subrate", "fdx_prospectradius", "fdx_prospectdatalastupdated");
                                Entity accountEntity = service.Retrieve("account", ((EntityReference)leadEntity.Attributes["parentaccountid"]).Id, columns);
                                accountid = accountEntity.Id;

                                //Case 1a: Lead when tagged with account having GM Accout Number
                                if (accountEntity.Attributes.Contains("fdx_goldmineaccountnumber"))
                                {
                                    #region Fetch account's address fields and update lead...

                                    step = 3;
                                    if (accountEntity.Attributes.Contains("fdx_goldmineaccountnumber"))
                                        leadEntity["fdx_goldmineaccountnumber"] = accountEntity.Attributes["fdx_goldmineaccountnumber"].ToString();

                                    step = 4;
                                    if (accountEntity.Attributes.Contains("fdx_gonogo"))
                                        leadEntity["fdx_gonogo"] = ((OptionSetValue)accountEntity.Attributes["fdx_gonogo"]).Value;

                                    step = 5;
                                    if (accountEntity.Attributes.Contains("fdx_zippostalcodeid"))
                                        leadEntity["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", ((EntityReference)accountEntity.Attributes["fdx_zippostalcodeid"]).Id);

                                    step = 6;
                                    //if (accountEntity.Attributes.Contains("telephone1"))
                                    //    leadEntity["telephone2"] = Regex.Replace(accountEntity.Attributes["telephone1"].ToString(),@"[^0-9]+", "");
                                    //leadEntity["telephone2"] = accountEntity.Attributes["telephone1"].ToString();

                                    step = 7;
                                    if (accountEntity.Attributes.Contains("name"))
                                        leadEntity["companyname"] = accountEntity.Attributes["name"];

                                    step = 8;
                                    if (accountEntity.Attributes.Contains("address1_line1"))
                                        leadEntity["address1_line1"] = accountEntity.Attributes["address1_line1"];

                                    step = 9;
                                    if (accountEntity.Attributes.Contains("address1_line2"))
                                        leadEntity["address1_line2"] = accountEntity.Attributes["address1_line2"];

                                    step = 10;
                                    if (accountEntity.Attributes.Contains("address1_city"))
                                        leadEntity["address1_city"] = accountEntity.Attributes["address1_city"];

                                    step = 11;
                                    if (accountEntity.Attributes.Contains("fdx_stateprovinceid"))
                                    {
                                        if (accountEntity.Attributes["fdx_stateprovinceid"] != null)
                                        {
                                            leadEntity["fdx_stateprovince"] = new EntityReference("fdx_state", ((EntityReference)accountEntity.Attributes["fdx_stateprovinceid"]).Id);
                                        }
                                    }

                                    step = 12;
                                    if (accountEntity.Attributes.Contains("address1_country"))
                                        leadEntity["address1_country"] = accountEntity.Attributes["address1_country"];

                                    ProspectData prospectData = GetProspectDataFromAccount(accountEntity);
                                    UpdateLeadWithImpersonation(leadEntity.Id, impersonatedService, prospectData);
                                    #endregion
                                }
                                //Case 1b: Lead when tagged with account which doesn't having GM Accout Number
                                else
                                {
                                    #region Fetch account's address fields and do create api call...
                                    updateAccount_GMNo = true;

                                    #region Call and update from API....
                                    step = 70;
                                    if (accountEntity.Attributes.Contains("fdx_zippostalcodeid"))
                                    {
                                        leadEntity["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", ((EntityReference)accountEntity.Attributes["fdx_zippostalcodeid"]).Id);
                                        apiParmAccCreate += string.Format("Zip={0}", (service.Retrieve("fdx_zipcode", ((EntityReference)accountEntity.Attributes["fdx_zippostalcodeid"]).Id, new ColumnSet("fdx_zipcode"))).Attributes["fdx_zipcode"].ToString());
                                    }
                                    step = 71;
                                    if (accountEntity.Attributes.Contains("telephone1"))
                                    {
                                        //leadEntity["telephone2"] = Regex.Replace(accountEntity.Attributes["telephone1"].ToString(),@"[^0-9]+", "");
                                        //leadEntity["telephone2"] = accountEntity.Attributes["telephone1"].ToString();
                                        apiParmAccCreate += string.Format("&Phone1={0}", Regex.Replace(accountEntity.Attributes["telephone1"].ToString(), @"[^0-9]+", ""));
                                        //apiParmAccCreate += string.Format("&Phone1={0}", accountEntity.Attributes["telephone1"].ToString());
                                    }

                                    step = 72;
                                    if (accountEntity.Attributes.Contains("name"))
                                    {
                                        leadEntity["companyname"] = accountEntity.Attributes["name"].ToString();
                                        apiParmAccCreate += string.Format("&Company={0}", accountEntity.Attributes["name"].ToString());
                                    }

                                    step = 73;
                                    if (accountEntity.Attributes.Contains("address1_line1"))
                                    {
                                        leadEntity["address1_line1"] = accountEntity.Attributes["address1_line1"].ToString();
                                        apiParmAccCreate += string.Format("&Address1={0}", accountEntity.Attributes["address1_line1"].ToString());
                                    }

                                    step = 74;
                                    if (accountEntity.Attributes.Contains("address1_line2"))
                                    {
                                        leadEntity["address1_line2"] = accountEntity.Attributes["address1_line2"].ToString();
                                        apiParmAccCreate += string.Format("&Address2={0}", accountEntity.Attributes["address1_line2"].ToString());
                                    }

                                    step = 75;
                                    if (accountEntity.Attributes.Contains("address1_city"))
                                    {
                                        leadEntity["address1_city"] = accountEntity.Attributes["address1_city"];
                                        apiParmAccCreate += string.Format("&City={0}", accountEntity.Attributes["address1_city"].ToString());
                                    }

                                    step = 76;
                                    if (accountEntity.Attributes.Contains("fdx_stateprovinceid"))
                                    {
                                        if (accountEntity.Attributes["fdx_stateprovinceid"] != null)
                                        {
                                            leadEntity["fdx_stateprovince"] = new EntityReference("fdx_state", ((EntityReference)accountEntity.Attributes["fdx_stateprovinceid"]).Id);
                                            apiParmAccCreate += string.Format("&State={0}", (service.Retrieve("fdx_state", ((EntityReference)accountEntity.Attributes["fdx_stateprovinceid"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString());
                                        }
                                    }

                                    step = 77;
                                    if (accountEntity.Attributes.Contains("address1_country"))
                                    {
                                        leadEntity["address1_country"] = accountEntity.Attributes["address1_country"].ToString();
                                        apiParmAccCreate += string.Format("&Country={0}", accountEntity.Attributes["address1_country"].ToString());
                                    }
                                    #endregion

                                    step = 78;

                                    url = smartCrmSyncWebServiceUrl + "/lead/createlead?" + apiParmAccCreate;

                                    createNewGMNo = true;
                                    #endregion
                                }
                                #endregion
                            }
                            //Case 2: Lead when tagged account is removed
                            else
                            {
                                step = 13;
                                #region Fetch lead's address fields and do create api call...

                                Entity ExistingLead = service.Retrieve("lead", leadEntity.Id, new ColumnSet("firstname", "lastname", "telephone2", "companyname", "address1_line1", "address1_city", "fdx_stateprovince", "fdx_zippostalcode", "address1_country", "fdx_jobtitlerole", "fdx_goldmineaccountnumber"));//, ));
                                #region create api param string...
                                //get lead data from query like account basedon lead id 

                                step = 14;
                                Guid zip = Guid.Empty;
                                if (leadEntity.Attributes.Contains("fdx_zippostalcode"))
                                    zip = ((EntityReference)leadEntity.Attributes["fdx_zippostalcode"]).Id;
                                else if (ExistingLead.Attributes.Contains("fdx_zippostalcode"))
                                    zip = ((EntityReference)ExistingLead.Attributes["fdx_zippostalcode"]).Id;

                                step = 15;
                                string zipcodetext = (service.Retrieve("fdx_zipcode", zip, new ColumnSet("fdx_zipcode"))).Attributes["fdx_zipcode"].ToString();
                                apiParmCreate += string.Format("Zip={0}", zipcodetext);

                                string firstName = ExistingLead.Attributes["firstname"].ToString();

                                step = 16;
                                string lastName = ExistingLead.Attributes["lastname"].ToString();

                                step = 17;
                                apiParmCreate += string.Format("&Contact={0} {1}", firstName, lastName);

                                string phone = Regex.Replace(ExistingLead.Attributes["telephone2"].ToString(), @"[^0-9]+", "");
                                //string phone = ExistingLead.Attributes["telephone2"].ToString();
                                apiParmCreate += string.Format("&Phone1={0}", phone);

                                step = 18;
                                string companyName = "";
                                if (ExistingLead.Attributes.Contains("companyname"))
                                {
                                    companyName = ExistingLead.Attributes["companyname"].ToString();
                                    apiParmCreate += string.Format("&Company={0}", companyName);
                                }

                                step = 19;
                                string title = "";
                                if (ExistingLead.Attributes.Contains("fdx_jobtitlerole"))
                                {
                                    title = CRMQueryExpression.GetOptionsSetTextForValue(service, "lead", "fdx_jobtitlerole", ((OptionSetValue)ExistingLead.Attributes["fdx_jobtitlerole"]).Value);
                                    apiParmCreate += string.Format("&Title={0}", title);
                                }

                                step = 20;
                                string address1 = "";
                                if (leadEntity.Attributes.Contains("address1_line1"))
                                {
                                    address1 = leadEntity.Attributes["address1_line1"].ToString();
                                    apiParmCreate += string.Format("&Address1={0}", address1);
                                }
                                else if (ExistingLead.Attributes.Contains("address1_line1"))
                                {
                                    address1 = ExistingLead.Attributes["address1_line1"].ToString();
                                    apiParmCreate += string.Format("&Address1={0}", address1);
                                }

                                step = 21;
                                string address2 = "";
                                if (leadEntity.Attributes.Contains("address1_line2"))
                                {
                                    address2 = leadEntity.Attributes["address1_line2"].ToString();
                                    apiParmCreate += string.Format("&Address2={0}", address2);
                                }
                                else if (ExistingLead.Attributes.Contains("address1_line2"))
                                {
                                    address2 = ExistingLead.Attributes["address1_line2"].ToString();
                                    apiParmCreate += string.Format("&Address2={0}", address2);
                                }

                                step = 22;
                                string city = "";
                                if (leadEntity.Attributes.Contains("address1_city"))
                                {
                                    city = leadEntity.Attributes["address1_city"].ToString();
                                    apiParmCreate += string.Format("&City={0}", city);
                                }
                                else if (ExistingLead.Attributes.Contains("address1_city"))
                                {
                                    city = ExistingLead.Attributes["address1_city"].ToString();
                                    apiParmCreate += string.Format("&City={0}", city);
                                }

                                step = 23;
                                string state = "";
                                if (leadEntity.Attributes.Contains("fdx_stateprovince"))
                                {
                                    if (leadEntity.Attributes["fdx_stateprovince"] != null)
                                    {
                                        state = (service.Retrieve("fdx_state", ((EntityReference)leadEntity.Attributes["fdx_stateprovince"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString();
                                        apiParmCreate += string.Format("&State={0}", state);
                                    }
                                }
                                else if (ExistingLead.Attributes.Contains("fdx_stateprovince"))
                                {
                                    if (ExistingLead.Attributes["fdx_stateprovince"] != null)
                                    {
                                        state = (service.Retrieve("fdx_state", ((EntityReference)ExistingLead.Attributes["fdx_stateprovince"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString();
                                        apiParmCreate += string.Format("&State={0}", state);
                                    }
                                }

                                step = 24;
                                string country = "";
                                if (leadEntity.Attributes.Contains("address1_country"))
                                {
                                    country = leadEntity.Attributes["address1_country"].ToString();
                                    apiParmCreate += string.Format("&Country={0}", country);
                                }
                                else if (ExistingLead.Attributes.Contains("address1_country"))
                                {
                                    country = ExistingLead.Attributes["address1_country"].ToString();
                                    apiParmCreate += string.Format("&Country={0}", country);
                                }
                                #endregion

                                step = 25;

                                url = smartCrmSyncWebServiceUrl + "/lead/createlead?" + apiParmAccCreate;

                                createNewGMNo = true;
                                step = 26;
                                #endregion
                            }
                        }
                        //Case 3: Lead with out any account tagged has a address change
                        if (step == 0)
                        {
                            #region Fetch lead's address fields and do update api call...

                            step = 51;
                            Entity ExistingLead = service.Retrieve("lead", leadEntity.Id, new ColumnSet("firstname", "lastname", "telephone2", "companyname", "address1_line1", "fdx_stateprovince", "address1_city", "fdx_zippostalcode", "address1_country", "fdx_jobtitlerole", "fdx_goldmineaccountnumber"));//, ));
                            #region create api param string...
                            //get lead data from query like account basedon lead id 

                            step = 59;
                            string goldmineaccno = "";
                            if (ExistingLead.Attributes.Contains("fdx_goldmineaccountnumber"))
                            {
                                goldmineaccno = ExistingLead.Attributes["fdx_goldmineaccountnumber"].ToString();
                                //Encoded the GM Account Number to overcome problem with Special Characters
                                string encodedgm = WebUtility.UrlEncode(ExistingLead.Attributes["fdx_goldmineaccountnumber"].ToString());
                                apiParmUpdate += string.Format("&AccountNo_in={0}", encodedgm);
                            }

                            step = 52;
                            Guid zip = Guid.Empty;
                            if (leadEntity.Attributes.Contains("fdx_zippostalcode"))
                                zip = ((EntityReference)leadEntity.Attributes["fdx_zippostalcode"]).Id;
                            else if (ExistingLead.Attributes.Contains("fdx_zippostalcode"))
                                zip = ((EntityReference)ExistingLead.Attributes["fdx_zippostalcode"]).Id;

                            step = 53;
                            string zipcodetext = (service.Retrieve("fdx_zipcode", zip, new ColumnSet("fdx_zipcode"))).Attributes["fdx_zipcode"].ToString();
                            apiParmUpdate += string.Format("&Zip={0}", zipcodetext);

                            step = 54;
                            string companyName = "";
                            if (ExistingLead.Attributes.Contains("companyname"))
                            {
                                companyName = ExistingLead.Attributes["companyname"].ToString();
                                apiParmUpdate += string.Format("&Company={0}", companyName);
                            }

                            step = 55;
                            string address1 = "";
                            if (leadEntity.Attributes.Contains("address1_line1"))
                            {
                                address1 = leadEntity.Attributes["address1_line1"].ToString();
                                apiParmUpdate += string.Format("&Address1={0}", address1);
                            }
                            else if (ExistingLead.Attributes.Contains("address1_line1"))
                            {
                                address1 = ExistingLead.Attributes["address1_line1"].ToString();
                                apiParmUpdate += string.Format("&Address1={0}", address1);
                            }

                            step = 56;
                            string city = "";
                            if (leadEntity.Attributes.Contains("address1_city"))
                            {
                                city = leadEntity.Attributes["address1_city"].ToString();
                                apiParmUpdate += string.Format("&City={0}", city);
                            }
                            else if (ExistingLead.Attributes.Contains("address1_city"))
                            {
                                city = ExistingLead.Attributes["address1_city"].ToString();
                                apiParmUpdate += string.Format("&City={0}", city);
                            }

                            step = 57;
                            string state = "";
                            if (leadEntity.Attributes.Contains("fdx_stateprovince"))
                            {
                                if (leadEntity.Attributes["fdx_stateprovince"] != null)
                                {
                                    state = (service.Retrieve("fdx_state", ((EntityReference)leadEntity.Attributes["fdx_stateprovince"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString();
                                    apiParmUpdate += string.Format("&State={0}", state);
                                }
                            }
                            else if (ExistingLead.Attributes.Contains("fdx_stateprovince"))
                            {
                                if (ExistingLead.Attributes["fdx_stateprovince"] != null)
                                {
                                    state = (service.Retrieve("fdx_state", ((EntityReference)ExistingLead.Attributes["fdx_stateprovince"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString();
                                    apiParmUpdate += string.Format("&State={0}", state);
                                }
                            }

                            step = 58;
                            string country = "";
                            if (leadEntity.Attributes.Contains("address1_country"))
                            {
                                country = leadEntity.Attributes["address1_country"].ToString();
                                apiParmUpdate += string.Format("&Country={0}", country);
                            }
                            else if (ExistingLead.Attributes.Contains("address1_country"))
                            {
                                country = ExistingLead.Attributes["address1_country"].ToString();
                                apiParmUpdate += string.Format("&Country={0}", country);
                            }


                            #endregion

                            step = 60;

                            url = smartCrmSyncWebServiceUrl + "/lead/updatelead?" + apiParmUpdate;

                            step = 61;
                            #endregion
                        }
                        tracingService.Trace("URL: {0}", url);
                        tracingService.Trace("createNewGMNo: {0}", createNewGMNo);
                        #region Call and update from API....
                        if (!string.IsNullOrEmpty(url))
                        {
                            step = 41;
                            const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
                            //This zipCode needs to be changed to that of Account
                            var uri = new Uri(url);
                            var request = WebRequest.Create(uri);
                            if (createNewGMNo)
                            {
                                request.Method = WebRequestMethods.Http.Post;
                            }
                            else
                            {
                                request.Method = WebRequestMethods.Http.Put;
                            }
                            request.ContentType = "application/json";
                            request.ContentLength = 0;
                            request.Headers.Add("Authorization", token);
                            step = 42;
                            using (var getResponse = request.GetResponse())
                            {
                                tracingService.Trace("createNewGMNo = " + createNewGMNo);
                                //This loop will be entered only if a new GM Account No is created only, and we will be calling POST method. And the response given by POST is serialised using Lead Class
                                if (createNewGMNo)
                                {
                                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Lead));
                                    Lead leadObj = new Lead();
                                    step = 43;
                                    leadObj = (Lead)serializer.ReadObject(getResponse.GetResponseStream());
                                    step = 44;
                                    leadEntity["fdx_goldmineaccountnumber"] = leadObj.goldMineId;

                                    if (leadObj.goNoGo)
                                    {
                                        step = 45;
                                        leadEntity["fdx_gonogo"] = new OptionSetValue(756480000);
                                    }
                                    else
                                    {
                                        step = 46;
                                        leadEntity["fdx_gonogo"] = new OptionSetValue(756480001);
                                    }
                                    
                                    ProspectData prospectData = GetProspectDataFromWebService(leadObj);
                                    PopulatePriceListIdAndProspectGroup(service, leadObj, prospectData);
                                    tracingService.Trace(GetProspectDataString(prospectData));
                                    UpdateLeadWithImpersonation(leadEntity.Id, impersonatedService, prospectData);
                                    tracingService.Trace("Prospect Data Updated");
                                    if (updateAccount_GMNo)
                                    {
                                        
                                        Entity accountUpdate = new Entity("account")
                                        {
                                            Id = accountid
                                        };
                                        accountUpdate.Attributes["fdx_goldmineaccountnumber"] = leadObj.goldMineId;
                                        accountUpdate.Attributes["fdx_gonogo"] = leadObj.goNoGo ? new OptionSetValue(756480000) : new OptionSetValue(756480001);
                                        UpdateProspectDataOnAccount(accountUpdate, prospectData);
                                        impersonatedService.Update(accountUpdate);
                                    }
                                }
                                //This loop will be entered only if Address is changed for an existing GM Account No, and we will be calling PUT method. And the response given by POST is serialised using API_PutResponse Class
                                else
                                {
                                    step = 47;
                                    DataContractJsonSerializer PutSerializer = new DataContractJsonSerializer(typeof(API_PutResponse));

                                    API_PutResponse LeadResponseObj = new API_PutResponse();
                                    LeadResponseObj = (API_PutResponse)PutSerializer.ReadObject(getResponse.GetResponseStream());
                                    if (LeadResponseObj.goNoGo)
                                    {
                                        step = 48;
                                        leadEntity["fdx_gonogo"] = new OptionSetValue(756480000);
                                    }
                                    else
                                    {
                                        step = 49;
                                        leadEntity["fdx_gonogo"] = new OptionSetValue(756480001);
                                    }
                                    ProspectData prospectData = GetProspectDataFromWebService(LeadResponseObj);
                                    PopulatePriceListIdAndProspectGroup(service, LeadResponseObj, prospectData);
                                    tracingService.Trace(GetProspectDataString(prospectData));
                                    UpdateLeadWithImpersonation(leadEntity.Id, impersonatedService, prospectData);
                                    tracingService.Trace("Prospect Data Updated");
                                }
                            }
                        }
                        step = 50;
                        #endregion
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    tracingService.Trace("LeadAddressUpdate: step {0}, {1}", step, ex.ToString());
                    throw new InvalidPluginExecutionException(string.Format("An error occurred in the LeadAddressUpdate plug-in at Step {0}.", step), ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("LeadAddressUpdate: step {0}, {1}", step, ex.ToString());
                    throw;
                }
            }
        }

        private void UpdateLeadWithImpersonation(Guid leadId, IOrganizationService impersonatedService, ProspectData prospectData)
        {
            UpdateRequest updateRequest = new UpdateRequest();
            Entity leadUpdateWithProspectingData = new Entity("lead", leadId);
            UpdateProspectData(leadUpdateWithProspectingData, prospectData);
            updateRequest.Target = leadUpdateWithProspectingData;
            updateRequest["SuppressDuplicateDetection"] = true;
            impersonatedService.Execute(updateRequest);
        }

        private void PopulatePriceListIdAndProspectGroup(IOrganizationService service, Lead leadObj, ProspectData prospectData)
        {
            EntityCollection priceLists = GetPriceListByName(leadObj.priceListName, service);
            EntityCollection prospectGroups = GetProspectGroupByName(leadObj.prospectGroup, service);
            if (priceLists.Entities.Count == 1)
                prospectData.PriceListId = priceLists.Entities[0].Id;
            if (prospectGroups.Entities.Count == 1)
                prospectData.ProspectGroupId = prospectGroups.Entities[0].Id;
        }

        private void PopulatePriceListIdAndProspectGroup(IOrganizationService service, API_PutResponse leadObj, ProspectData prospectData)
        {
            EntityCollection priceLists = GetPriceListByName(leadObj.priceListName, service);
            EntityCollection prospectGroups = GetProspectGroupByName(leadObj.prospectGroup, service);
            if (priceLists.Entities.Count == 1)
                prospectData.PriceListId = priceLists.Entities[0].Id;
            if (prospectGroups.Entities.Count == 1)
                prospectData.ProspectGroupId = prospectGroups.Entities[0].Id;
        }

        private ProspectData GetProspectData()
        {
            ProspectData prospectData = new ProspectData();
            prospectData.ProspectGroupId = new Guid("9B3945FC-2728-E811-811D-3863BB34CB20");
            prospectData.PriceListId = new Guid("8A826A97-0B26-E811-811C-3863BB35EF70");
            prospectData.Priority = Convert.ToDecimal(1);
            prospectData.Score = Convert.ToDecimal(2);
            prospectData.Percentile = Convert.ToDecimal(3);
            prospectData.RateSource = "Stub";
            prospectData.PPRRate = Convert.ToDecimal(1);
            prospectData.SubRate = Convert.ToDecimal(2);
            prospectData.Radius = 2;
            return prospectData;
        }

        private ProspectData GetProspectDataFromWebService(API_PutResponse lead)
        {
            ProspectData prospectData = new ProspectData();
            prospectData.ProspectGroupName = lead.prospectGroup;
            prospectData.PriceListName = lead.priceListName;
            prospectData.Priority = lead.prospectPriority;
            prospectData.Score = lead.prspectScore;
            prospectData.Percentile = lead.prospectPercentile;
            prospectData.RateSource = lead.rateSource;
            prospectData.PPRRate = lead.pprRate;
            prospectData.SubRate = lead.subRate;
            prospectData.Radius = lead.prospectRadius;
            return prospectData;
        }

        private ProspectData GetProspectDataFromWebService(Lead lead)
        {
            ProspectData prospectData = new ProspectData();
            prospectData.ProspectGroupName = lead.prospectGroup;
            prospectData.PriceListName = lead.priceListName;
            prospectData.Priority = lead.prospectPriority;
            prospectData.Score = lead.prspectScore;
            prospectData.Percentile = lead.prospectPercentile;
            prospectData.RateSource = lead.rateSource;
            prospectData.PPRRate = lead.pprRate;
            prospectData.SubRate = lead.subRate;
            prospectData.Radius = lead.prospectRadius;
            return prospectData;
        }

        private ProspectData GetProspectDataFromAccount(Entity account)
        {
            ProspectData prospectData = new ProspectData();
            if (account.Contains("fdx_prospectgroup"))
                prospectData.ProspectGroupId = ((EntityReference)account["fdx_prospectgroup"]).Id;
            if (account.Contains("defaultpricelevelid"))
            {
                EntityReference priceList = (EntityReference)account["defaultpricelevelid"];
                prospectData.PriceListId = priceList.Id;
                prospectData.PriceListName = priceList.Name;
            }
            if (account.Contains("fdx_prospectpriority"))
                prospectData.Priority = (decimal)account["fdx_prospectpriority"];
            if (account.Contains("fdx_prospectscore"))
                prospectData.Score = (decimal)account["fdx_prospectscore"];
            if (account.Contains("fdx_prospectpercentile"))
                prospectData.Percentile = (decimal)account["fdx_prospectpercentile"];
            if (account.Contains("fdx_ratesource"))
                prospectData.RateSource = (string)account["fdx_ratesource"];
            if (account.Contains("fdx_pprrate"))
                prospectData.PPRRate = ((Money)account["fdx_pprrate"]).Value;
            if (account.Contains("fdx_subrate"))
                prospectData.SubRate = ((Money)account["fdx_subrate"]).Value;
            if (account.Contains("fdx_prospectradius"))
                prospectData.Radius = (int)account["fdx_prospectradius"];
            if (account.Contains("fdx_prospectdatalastupdated"))
                prospectData.LastUpdated = (DateTime)account["fdx_prospectdatalastupdated"];
            return prospectData;
        }

        private string GetProspectDataString(ProspectData prospectData)
        {
            string traceString = "ProspectGroupName=" + prospectData.ProspectGroupName + Environment.NewLine;
            traceString += "PriceListName=" + prospectData.PriceListName + Environment.NewLine;
            traceString += "Priority=" + prospectData.Priority.ToString() + Environment.NewLine;
            traceString += "Score=" + prospectData.Score.ToString() + Environment.NewLine;
            traceString += "Percentile=" + prospectData.Percentile.ToString() + Environment.NewLine;
            traceString += "RateSource=" + prospectData.RateSource + Environment.NewLine;
            traceString += "PPRRate=" + prospectData.PPRRate.ToString() + Environment.NewLine;
            traceString += "SubRate=" + prospectData.SubRate.ToString() + Environment.NewLine;
            traceString += "Radius=" + prospectData.Radius.ToString() + Environment.NewLine;
            return traceString;
        }

        private void UpdateProspectData(Entity leadRecord, ProspectData prospectData)
        {
            if(prospectData.ProspectGroupId.HasValue)
            {
                leadRecord["fdx_prospectgroup"] =  new EntityReference("fdx_prospectgroup", prospectData.ProspectGroupId.Value);
            }
            else
            {
                leadRecord["fdx_prospectgroup"] = null;
            }

            if (prospectData.PriceListId.HasValue)
            {
                leadRecord["fdx_pricelist"] = new EntityReference("pricelevel", prospectData.PriceListId.Value);
            }
            else
            {
                leadRecord["fdx_pricelist"] = null;
                leadRecord["fdx_prospectpricelistname"] = null;
            }
            
            leadRecord["fdx_prospectpriority"] = prospectData.Priority.HasValue ? prospectData.Priority : null;
            leadRecord["fdx_prospectscore"] = prospectData.Score.HasValue ? prospectData.Score : null;
            leadRecord["fdx_prospectpercentile"] =prospectData.Percentile.HasValue ? prospectData.Percentile : null;
            leadRecord["fdx_ratesource"] = prospectData.RateSource;
            leadRecord["fdx_pprrate"] =prospectData.PPRRate.HasValue ? new Money(prospectData.PPRRate.Value) : null;
            leadRecord["fdx_subrate"] =prospectData.SubRate.HasValue ? new Money(prospectData.SubRate.Value) : null;
            leadRecord["fdx_prospectradius"] = prospectData.Radius.HasValue? prospectData.Radius : null;
            if (prospectData.LastUpdated.HasValue)
            {
                leadRecord["fdx_prospectdatalastupdated"] = prospectData.LastUpdated.Value;
            }
            else
            {
                leadRecord["fdx_prospectdatalastupdated"] = DateTime.UtcNow;
            }
        }

        private void UpdateProspectDataOnAccount(Entity accountRecord, ProspectData prospectData)
        {
            if (prospectData.ProspectGroupId.HasValue && !prospectData.ProspectGroupId.Equals(Guid.Empty))
                accountRecord["fdx_prospectgroup"] = new EntityReference("fdx_prospectgroup", prospectData.ProspectGroupId.Value);
            if (prospectData.PriceListId.HasValue && !prospectData.PriceListId.Equals(Guid.Empty))
                accountRecord["defaultpricelevelid"] = new EntityReference("pricelevel", prospectData.PriceListId.Value);
            if (prospectData.Priority.HasValue)
                accountRecord["fdx_prospectpriority"] = prospectData.Priority;
            if (prospectData.Score.HasValue)
                accountRecord["fdx_prospectscore"] = prospectData.Score;
            if (prospectData.Percentile.HasValue)
                accountRecord["fdx_prospectpercentile"] = prospectData.Percentile;
            if (!string.IsNullOrEmpty(prospectData.RateSource))
                accountRecord["fdx_ratesource"] = prospectData.RateSource;
            if (prospectData.PPRRate.HasValue)
                accountRecord["fdx_pprrate"] = new Money(prospectData.PPRRate.Value);
            if (prospectData.SubRate.HasValue)
                accountRecord["fdx_subrate"] = new Money(prospectData.SubRate.Value);
            if (prospectData.Radius.HasValue)
                accountRecord["fdx_prospectradius"] = prospectData.Radius;
            accountRecord["fdx_prospectdatalastupdated"] = DateTime.UtcNow;
        }

        private EntityCollection GetPriceListByName(string priceListName, IOrganizationService crmService)
        {
            QueryByAttribute queryByPriceList = new QueryByAttribute("pricelevel");
            queryByPriceList.ColumnSet = new ColumnSet("pricelevelid");
            queryByPriceList.AddAttributeValue("name", priceListName);
            return crmService.RetrieveMultiple(queryByPriceList);
        }

        private EntityCollection GetProspectGroupByName(string prospectGroupName, IOrganizationService crmService)
        {
            QueryByAttribute queryByProspectGroup = new QueryByAttribute("fdx_prospectgroup");
            queryByProspectGroup.ColumnSet = new ColumnSet("fdx_prospectgroupid");
            queryByProspectGroup.AddAttributeValue("fdx_name", prospectGroupName);
            return crmService.RetrieveMultiple(queryByProspectGroup);
        }
    }
}