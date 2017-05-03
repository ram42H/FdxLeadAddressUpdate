using FdxLeadAssignmentPlugin;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace FdxLeadAddressUpdate
{
    public class LeadAddressUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
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

            //Call Input parameter collection to get all the data passes....
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity leadEntity = (Entity)context.InputParameters["Target"];

                if (leadEntity.LogicalName != "lead")
                    return;

                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    step = 0;
                    Guid accountid = Guid.Empty;

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
                                Entity accountEntity = service.Retrieve("account", ((EntityReference)leadEntity.Attributes["parentaccountid"]).Id, new ColumnSet("name", "fdx_goldmineaccountnumber", "fdx_gonogo", "address1_line1", "address1_line2", "address1_city", "fdx_stateprovinceid", "fdx_zippostalcodeid", "telephone1", "address1_country"));
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
                                    if (accountEntity.Attributes.Contains("telephone1"))
                                        leadEntity["telephone2"] = accountEntity.Attributes["telephone1"];
                                    
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
                                        leadEntity["telephone2"] = accountEntity.Attributes["telephone1"].ToString();
                                        apiParmAccCreate += string.Format("&Phone1={0}", accountEntity.Attributes["telephone1"].ToString());
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
                                    url = "http://SMARTCRMSync.1800dentist.com/api/lead/createleadasync?" + apiParmAccCreate;
                                    #endregion
                                }
                                #endregion
                            }
                            //Case 2: Lead when tagged account is removed
                            else
                            {
                                step = 13;
                                #region Fetch lead's address fields and do create api call...

                                Entity ExistingLead = service.Retrieve("lead", leadEntity.Id, new ColumnSet("firstname", "lastname", "telephone2", "companyname", "address1_line1","address1_city", "fdx_stateprovince", "fdx_zippostalcode", "address1_country", "fdx_jobtitlerole", "fdx_goldmineaccountnumber"));//, ));
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

                                string phone = ExistingLead.Attributes["telephone2"].ToString();
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
                                url = "http://SMARTCRMSync.1800dentist.com/api/lead/createleadasync?" + apiParmCreate;
                                step = 26;
                                #endregion
                            }
                        }
                        //Case 3: Lead with out any account tagged has a address change
                        if (step == 0) 
                        {
                            #region Fetch lead's address fields and do update api call...

                            step = 51;
                            Entity ExistingLead = service.Retrieve("lead", leadEntity.Id, new ColumnSet("firstname", "lastname", "telephone2", "companyname", "address1_line1", "fdx_stateprovince","address1_city", "fdx_zippostalcode", "address1_country", "fdx_jobtitlerole", "fdx_goldmineaccountnumber"));//, ));
                            #region create api param string...
                            //get lead data from query like account basedon lead id 

                            step = 59;
                            string goldmineaccno = "";
                            if (ExistingLead.Attributes.Contains("fdx_goldmineaccountnumber"))
                            {
                                goldmineaccno = ExistingLead.Attributes["fdx_goldmineaccountnumber"].ToString();
                                apiParmUpdate += string.Format("&AccountNo_in={0}", ExistingLead.Attributes["fdx_goldmineaccountnumber"].ToString());
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
                            url = "http://smartcrmsync.1800dentist.com/api/lead/updateleadasync?" + apiParmUpdate;
                            step = 61;
                            #endregion
                        }

                        #region Call and update from API....
                        if (!string.IsNullOrEmpty(url))
                        {
                            Lead leadObj = new Lead();
                            step = 41;
                            const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
                            //This zipCode needs to be changed to that of Account
                            var uri = new Uri(url);
                            var request = WebRequest.Create(uri);
                            request.Method = WebRequestMethods.Http.Post;
                            request.ContentType = "application/json";
                            request.ContentLength = 0;
                            request.Headers.Add("Authorization", token);
                            step = 42;
                            using (var getResponse = request.GetResponse())
                            {
                                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Lead));

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
                                if (updateAccount_GMNo)
                                {
                                    Entity accountUpdate = new Entity("account")
                                    {
                                        Id = accountid
                                    };
                                    accountUpdate.Attributes["fdx_goldmineaccountnumber"] = leadObj.goldMineId;
                                    accountUpdate.Attributes["fdx_gonogo"] = leadObj.goNoGo ? new OptionSetValue(756480000) : new OptionSetValue(756480001);
                                    service.Update(accountUpdate);
                                }
                            }
                        }
                        step = 47;
                        #endregion
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("An error occurred in the LeadAddressUpdate plug-in at Step {0}.", step), ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("LeadAddressUpdate: step {0}, {1}", step, ex.ToString());
                    throw;
                }
            }
        }
    }
}
