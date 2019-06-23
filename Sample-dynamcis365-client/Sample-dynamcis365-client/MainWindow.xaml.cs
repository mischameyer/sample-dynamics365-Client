using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Configuration;
// These namespaces are found in the Microsoft.Crm.Sdk.Proxy.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Crm.Sdk.Messages;

// These namespaces are found in the Microsoft.Xrm.Sdk.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace Sample_dynamcis365_client
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Class Level Members

        private Guid _accountId;
        private IOrganizationService _orgService;

        #endregion Class Level Members

        public MainWindow()
        {
            InitializeComponent();

            //disable the action buttons when until a connection is available
            this.btnCreate.IsEnabled = false;
            this.btnRetrieve.IsEnabled = false;
            this.btnUpdate.IsEnabled = false;
            this.btnDelete.IsEnabled = false;

            // Read the server configurations from app.config.
            GetServiceConfiguration();
        }

        #region PublicMethods
        /// <summary>
        /// The Connect() method first connects to the organization service. 
        /// </summary>
        /// <param name="connectionString">Provides service connection information.</param>
        /// <param name="promptforDelete">When True, the user will be prompted to delete all
        /// created entities.</param>
        public void Connect(String connectionString, bool promptforDelete)
        {
            try
            {
                // Establish a connection to the organization web service.
                Print("Connecting to the server ...");

                // Connect to the CRM web service using a connection string.
                CrmServiceClient conn = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(connectionString);

                // Cast the proxy client to the IOrganizationService interface.
                _orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                Print("connected");

                // Obtain information about the logged on user from the web service.
                Guid userid = ((WhoAmIResponse)_orgService.Execute(new WhoAmIRequest())).UserId;
                SystemUser systemUser = (SystemUser)_orgService.Retrieve("systemuser", userid,
                    new ColumnSet(new string[] { "firstname", "lastname" }));
                Println("Logged on user is " + systemUser.FirstName + " " + systemUser.LastName + ".");

                // Retrieve the version of Microsoft Dynamics CRM.
                RetrieveVersionRequest versionRequest = new RetrieveVersionRequest();
                RetrieveVersionResponse versionResponse =
                    (RetrieveVersionResponse)_orgService.Execute(versionRequest);
                Println("Microsoft Dynamics CRM version " + versionResponse.Version + ".");

                //enable the action buttons when a connection is available
                this.btnCreate.IsEnabled = true;
                this.btnRetrieve.IsEnabled = true;
                this.btnUpdate.IsEnabled = true;
                this.btnDelete.IsEnabled = true;

            }

            // Catch any service fault exceptions that Microsoft Dynamics CRM throws.
            catch (FaultException<OrganizationServiceFault>)
            {
                // You can handle an exception here or pass it back to the calling method.
                throw;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void CreateAccount()
        {

            // Instantiate an account object. Note the use of option set enumerations defined in OptionSets.cs.
            // Refer to the Entity Metadata topic in the SDK documentation to determine which attributes must
            // be set for each entity.
            Account account = new Account { Name = this.txtAccountName.Text };
            account.AccountCategoryCode = new OptionSetValue((int)AccountAccountCategoryCode.PreferredCustomer);
            account.CustomerTypeCode = new OptionSetValue((int)AccountCustomerTypeCode.Investor);

            // Create an account record
            _accountId = _orgService.Create(account);

            Println(account.LogicalName + " " + account.Name + " created, ");

        }

        /// <summary>
        /// Retrieves an existing Account
        /// </summary>
        /// <returns></returns>
        public Account RetrieveAccount()
        {
            // Retrieve several attributes from the new account.
            ColumnSet cols = new ColumnSet(
                new String[] { "name", "address1_postalcode", "lastusedincampaign" });

            Account retrievedAccount = (Account)_orgService.Retrieve("account", _accountId, cols);

            Println("name: " + retrievedAccount.Name + ", adress1_postalcode: " + retrievedAccount.Address1_PostalCode + ", lastusedincampaign: " + retrievedAccount.LastUsedInCampaign);

            Println("Account retrieved...");

            return retrievedAccount;
        }

        /// <summary>
        /// Updates an existing Account
        /// </summary>
        /// <param name="retrievedAccount"></param>
        public void UpdateAccount(Account retrievedAccount)
        {
            // Update the postal code attribute.
            retrievedAccount.Address1_PostalCode = "98052";

            // The address 2 postal code was set accidentally, so set it to null.
            retrievedAccount.Address2_PostalCode = null;

            // Shows use of a Money value.
            retrievedAccount.Revenue = new Money(5000000);

            // Shows use of a Boolean value.
            retrievedAccount.CreditOnHold = false;

            // Update the account record.
            _orgService.Update(retrievedAccount);
            Println("Account updated.");
        }

        /// <summary>
        /// Deletes any entity records that were created for this sample.
        /// <param name="prompt">Indicates whether to prompt the user 
        /// to delete the records created in this sample.</param>
        /// </summary>
        public void DeleteRequiredRecords()
        {
            _orgService.Delete(Account.EntityLogicalName, _accountId);
            _accountId = Guid.Empty;
            Println("Entity records have been deleted.");
        }

        /// <summary>
        /// Displays a message string in the form with newline.
        /// </summary>
        public void Println(string _sPrintlntext)
        {

            if (lblOutMsg.Text != string.Empty)
                lblOutMsg.Text = lblOutMsg.Text + "\n" + _sPrintlntext;
            else
                lblOutMsg.Text = _sPrintlntext;

        }

        /// <summary>
        /// Displays a message string in the form.   
        /// </summary>
        public void Print(string _sPrintlntext)
        {
            lblOutMsg.Text = lblOutMsg.Text + _sPrintlntext;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Gets web service connection information from the app.config file.
        /// If there is more than one available, providing an option to the user to select
        /// the desired connection configuration by name.
        /// </summary>
        /// <returns>A string containing web service connection configuration information.</returns>
        private String GetServiceConfiguration()
        {
            // Get available connection strings from app.config.
            int count = ConfigurationManager.ConnectionStrings.Count;
            // Get the ConnectionStrings collection.

            // Create a filter list of connection strings so that we have a list of valid
            // connection strings for Microsoft Dynamics CRM only.
            List<KeyValuePair<String, String>> filteredConnectionStrings =
                new List<KeyValuePair<String, String>>();

            for (int a = 0; a < count; a++)
            {
                if (isValidConnectionString(ConfigurationManager.ConnectionStrings[a].ConnectionString))
                    filteredConnectionStrings.Add
                        (new KeyValuePair<string, string>
                            (ConfigurationManager.ConnectionStrings[a].Name,
                            ConfigurationManager.ConnectionStrings[a].ConnectionString));
            }

            // No valid connections strings found. Write out an error message.
            if (filteredConnectionStrings.Count == 0)
            {
                Println("An app.config file containing at least one valid Microsoft Dynamics CRM " +
                    "server connection configuration must exist in the run-time folder.");
                Println("\nThere are several commented out example server connection configurations in " +
                    "the provided app.config file. Uncomment one or more of them, modify the configuration according " +
                    "to your Microsoft Dynamics CRM installation, and then re-run the sample.");

                // Disable the Connect button.
                btnConnect.IsEnabled = false;
                return null;
            }

            // If at least one valid connection string is found, display the list of valid connection strings.
            else
            {
                for (int i = 0; i < filteredConnectionStrings.Count; i++)
                {
                    cbxServerList.Items.Add(filteredConnectionStrings[i].Key);

                }
                cbxServerList.SelectedIndex = 0;
            }

            // Return a non-null which in this case is the first string in the list. 
            return ConfigurationManager.ConnectionStrings[0].ConnectionString;
        }


        /// <summary>
        /// Verifies if a connection string is valid for Microsoft Dynamics CRM.
        /// </summary>
        /// <returns>True for a valid string, otherwise False.</returns>
        private static Boolean isValidConnectionString(String connectionString)
        {
            // At a minimum, a connection string must contain one of these arguments.
            if (connectionString.Contains("Url=") ||
                connectionString.Contains("Server=") ||
                connectionString.Contains("ServiceUri="))
                return true;

            return false;
        }


        /// <summary>
        /// Let the user choose which connection string to use.
        /// Gets the user selected  web service connection  from the app.config file.      
        /// </summary>
        /// <returns></returns>      
        private void btnConnect_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;

            try
            {
                string connectionString =
                    ConfigurationManager.ConnectionStrings[cbxServerList.SelectedItem.ToString()].ConnectionString;

                lblOutMsg.Text = string.Empty;

                if (connectionString != null) Connect(connectionString, true);
            }

            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Println("The application terminated with an error.");
                Println("Timestamp: " + ex.Detail.Timestamp);
                Println("Code: " + ex.Detail.ErrorCode);
                Println("Message: " + ex.Detail.Message);
                Println("Trace: " + ex.Detail.TraceText);
                Println("Inner Fault: {0}" + (null == ex.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault"));
            }
            catch (System.TimeoutException ex)
            {
                Println("The application terminated with an error.");
                Println("Message: " + ex.Message);
                Println("Stack Trace: " + ex.StackTrace);
                Println("Inner Fault: " + (null == ex.InnerException.Message ? "No Inner Fault" : ex.InnerException.Message));
            }
            catch (System.Exception ex)
            {
                Println("The application terminated with an error.");
                Println(ex.Message);

                // Display the details of the inner exception.
                if (ex.InnerException != null)
                {
                    Println(ex.InnerException.Message);

                    FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> fe = ex.InnerException
                        as FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>;
                    if (fe != null)
                    {
                        Println("Timestamp: " + fe.Detail.Timestamp);
                        Println("Code: " + fe.Detail.ErrorCode);
                        Println("Message: " + fe.Detail.Message);
                        Println("Trace: " + fe.Detail.TraceText);
                        Println("Inner Fault: " + (null == fe.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault"));
                    }
                }
            }

            // Additional exceptions to catch: SecurityTokenValidationException, ExpiredSecurityTokenException,
            // SecurityAccessDeniedException, MessageSecurityException, and SecurityNegotiationException.

            finally
            {
                Println("Let's try some CRUD operations...");
                this.Cursor = null;
            }
        }

        private void BtnCreate_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            CreateAccount();
            this.Cursor = null;
        }

        private void BtnRetrieve_Click(object sender, RoutedEventArgs e)
        {
            if (_accountId != Guid.Empty)
            {
                this.Cursor = Cursors.Wait;
                RetrieveAccount();
                this.Cursor = null;
            }
            else
            {
                Println("No Account set! Please create an account.");
            }
        }

        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (_accountId != Guid.Empty)
            {
                this.Cursor = Cursors.Wait;
                UpdateAccount(RetrieveAccount());
                this.Cursor = null;
            }
            else
            {
                Println("No Account set! Please create an account.");
            }

        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (_accountId != Guid.Empty)
            {
                this.Cursor = Cursors.Wait;
                DeleteRequiredRecords();
                this.Cursor = null;
            }
            else
            {
                Println("No Account set! Please create an account.");
            }
        }

        #endregion Private Methods            

    }
}
