using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TAPI3Lib;

namespace CustomAlertBoxDemo
{
    public partial class Form1 : Form
    {
		public static string domainname = ConfigurationManager.AppSettings.Get("adDomain");
		public static string domainUser = ConfigurationManager.AppSettings.Get("adUser");
		public static string domainPass = ConfigurationManager.AppSettings.Get("adPassword");
		//public static string adContainer = ConfigurationManager.AppSettings.Get("adContainer");
		//public static ContextType contexDom = ContextType.Domain;
		private static int index;
		public static ContextOptions contexOpt = (ContextOptions.SecureSocketLayer | ContextOptions.Negotiate);
		private StreamWriter errorLogYaz = new StreamWriter("ErrorLog.txt", true);
		private TAPIClass tobj;
		private ITAddress[] ia = new TAPI3Lib.ITAddress[10];
		private ITBasicCallControl bcc;
		//private callnotification cn;
		private bool h323, reject;
		uint lines;
		int line;
		int[] registertoken = new int[10];
		public Form1()
        {
            InitializeComponent();
			//MessageBox.Show("s", "");
			initializetapi3();
			h323 = false;
			reject = false;
		}

        public void Alert(string msg, Form_Alert.enmType type)
        {
            Form_Alert frm = new Form_Alert();
            frm.showAlert(msg,type);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Alert("Success Alert",Form_Alert.enmType.Success);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Alert("Warning Alert", Form_Alert.enmType.Warning);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Alert("Error Alert", Form_Alert.enmType.Error);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Alert("Info Alert", Form_Alert.enmType.Info);
        }

		void initializetapi3()
		{
			try
			{
				string telephoneNumber = this.getNumber(Environment.UserName);
				tobj = new TAPIClass();
				tobj.Initialize();
				IEnumAddress ea = tobj.EnumerateAddresses();
				ITAddress ln;
				uint arg3 = 0;
				lines = 0;
				this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString() + " -Info- " + "Line searcher starting here..." );
				//cn = new callnotification();
				//cn.addtolist = new callnotification.listshow(this.status);
				tobj.ITTAPIEventNotification_Event_Event += new TAPI3Lib.ITTAPIEventNotification_EventEventHandler(Event);
				tobj.EventFilter = (int)(TAPI_EVENT.TE_CALLNOTIFICATION |
					TAPI_EVENT.TE_DIGITEVENT |
					TAPI_EVENT.TE_PHONEEVENT |
					TAPI_EVENT.TE_CALLSTATE |
					TAPI_EVENT.TE_GENERATEEVENT |
					TAPI_EVENT.TE_GATHERDIGITS |
					TAPI_EVENT.TE_REQUEST);
				
				for (int i = 0; i < 10; i++)
				{
					ea.Next(1, out ln, ref arg3);
					ia[i] = ln;
					if (ln != null)
					{
						//if (ln.DialableAddress == telephoneNumber)
						//{
						//	index = i;
						//	this.registerByAD(index);
						//}
						this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString() + " -Info- " + "Line:"+i+" AdressName" + ia[i].AddressName);
						
						if (ia[i].AddressName.Contains("3000"))
							this.registertoken[2] = this.tobj.RegisterCallNotifications(ia[2], true, true, 8, 2);

						textBox2.Text = this.registertoken[2].ToString();
						//this.registerByAD(i);
						//comboBox1.Items.Add(ia[i].AddressName);
						lines++;
					}
					else
						break;
				}
				this.errorLogYaz.WriteLine("\n"+ DateTime.Now.ToString() + " -Info- " + "Line searcher ending here...");
				//tobj.ITTAPIEventNotification_Event_Event+= new TAPI3Lib.ITTAPIEventNotification_EventEventHandler(cn.Event);
				//tobj.EventFilter=(int)(TAPI_EVENT.TE_CALLNOTIFICATION|TAPI_EVENT.TE_DIGITEVENT|TAPI_EVENT.TE_PHONEEVENT|TAPI_EVENT.TE_CALLSTATE);
				//registertoken=tobj.RegisterCallNotifications(ia[6],true,true,TapiConstants.TAPIMEDIATYPE_AUDIO|TapiConstants.TAPIMEDIATYPE_DATAMODEM,1);	
				//MessageBox.Show("Registration token :-"+registertoken,"Regitration complete");

			}
			catch (Exception e)
			{
				MessageBox.Show(e.ToString());
				this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString() + " -Error- " + e.ToString());
			}
		}
		public void registerByAD(int line)
		{
			try
			{
				this.registertoken[line] = this.tobj.RegisterCallNotifications(ia[line], true, true, 8, 2);
				string username = this.getUserName(Environment.UserName);
				string telephoneNumber = this.getNumber(Environment.UserName);
				MessageBox.Show("Success to register on line " + ((int)line).ToString(), "Username:"+ username + " Number:"+ telephoneNumber);
			}
			catch (Exception)
			{
				MessageBox.Show("Failed to register on line " + ((int)line).ToString(), "Registration for calls");
				this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString()+" -Error- "+"Failed to register on line " + ((int)line).ToString(), "Registration for calls");
			}
		}
		private string getNumber(string username)
		{
			string voiceTelephoneNumber;
			try
			{
				PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, domainname, domainUser, domainPass);
				UserPrincipal userPrincipal = new UserPrincipal(principalContext);
				userPrincipal.Enabled = true;
				PrincipalSearcher principalSearcher = new PrincipalSearcher(userPrincipal);
				List<UserPrincipal> principalSearchResult = principalSearcher.FindAll().Cast<UserPrincipal>().OrderBy(u => u.SamAccountName).ToList();

				/*using (PrincipalContext context = new PrincipalContext(contexDom, domainname, adContainer, contexOpt, domainUser, domainPass))*/
				{
					UserPrincipal queryFilter = new UserPrincipal(principalContext)
					{
						SamAccountName = username
					};
					PrincipalSearcher searcher = new PrincipalSearcher(queryFilter);
					try
					{
						UserPrincipal principal2 = (UserPrincipal)searcher.FindOne();
						if (principal2 != null)
						{
							voiceTelephoneNumber = principal2.VoiceTelephoneNumber;
						}
						else
						{
							this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString() + " -Error- "+"Active directory'de bu kullanıcıya ait bir numara yok."  );
							MessageBox.Show("Bu kullanıcıya ait bir numara yok.");
							voiceTelephoneNumber = null;
						}
					}
					catch (Exception exception)
					{
						this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString()+" -Error- " +exception.ToString() );
						voiceTelephoneNumber = null;
					}
				}
			}
			catch (Exception exception)
			{
				voiceTelephoneNumber = null;
				this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString() + " -Error- " + exception.ToString());
			}
			return voiceTelephoneNumber;
		}
		private string getUserName(string username)
		{
			string displayName;

			PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, domainname, domainUser, domainPass);
			UserPrincipal userPrincipal = new UserPrincipal(principalContext);
			userPrincipal.Enabled = true;
			PrincipalSearcher principalSearcher = new PrincipalSearcher(userPrincipal);
			List<UserPrincipal> principalSearchResult = principalSearcher.FindAll().Cast<UserPrincipal>().OrderBy(u => u.SamAccountName).ToList();

			//using (PrincipalContext context = new PrincipalContext(contexDom, domainname, adContainer, contexOpt, domainUser, domainPass))
			{
				UserPrincipal queryFilter = new UserPrincipal(principalContext)
				{
					SamAccountName = username
				};
				PrincipalSearcher searcher = new PrincipalSearcher(queryFilter);
				try
				{
					UserPrincipal principal2 = (UserPrincipal)searcher.FindOne();
					if (principal2 != null)
					{
						displayName = principal2.DisplayName;
					}
					else
					{
						MessageBox.Show("Bu kullanıcı tanımlı değil");
						this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString() + " -Error- " + principal2 + "Kullanıcı Tanımlı değil");
						displayName = null;
					}
				}
				catch (Exception exp)
				{
					this.errorLogYaz.WriteLine("\n" + DateTime.Now.ToString() + " -Error- " + exp.ToString());
					MessageBox.Show("Numara girin");
					displayName = null;
				}
			}
			return displayName;
		}
		public void Event(TAPI3Lib.TAPI_EVENT te, object eobj)
		{
			ITCallStateEvent event4 = (ITCallStateEvent)eobj;
			ITCallNotificationEvent event5 = eobj as ITCallNotificationEvent;
			string cnumber = event4.Call.get_CallInfoString(CALLINFO_STRING.CIS_CALLERIDNUMBER);
			switch (te)
			{
				case TAPI3Lib.TAPI_EVENT.TE_CALLNOTIFICATION:
					this.Alert("call notification event has occured", Form_Alert.enmType.Info);					
					break;
				case TAPI3Lib.TAPI_EVENT.TE_DIGITEVENT:
					TAPI3Lib.ITDigitDetectionEvent dd = (TAPI3Lib.ITDigitDetectionEvent)eobj;
					this.Alert("Dialed digit" + dd.ToString(), Form_Alert.enmType.Info);
					break;
				case TAPI3Lib.TAPI_EVENT.TE_GENERATEEVENT:
					TAPI3Lib.ITDigitGenerationEvent dg = (TAPI3Lib.ITDigitGenerationEvent)eobj;
					this.Alert("Dialed digit" + dg.ToString(), Form_Alert.enmType.Info);
					break;
				case TAPI3Lib.TAPI_EVENT.TE_PHONEEVENT:
					this.Alert("A phone event!", Form_Alert.enmType.Info);
					break;
				case TAPI3Lib.TAPI_EVENT.TE_GATHERDIGITS:
					this.Alert("Gather digit event!", Form_Alert.enmType.Info);
					break;
				case TAPI3Lib.TAPI_EVENT.TE_CALLSTATE:
					TAPI3Lib.ITCallStateEvent a = (TAPI3Lib.ITCallStateEvent)eobj;
					TAPI3Lib.ITCallInfo b = a.Call;

					/*if (b.CallState== TAPI3Lib.CALL_STATE.CS_INPROGRESS)
                    {
						MessageBox.Show("dialing");
						this.Alert("dialing", Form_Alert.enmType.Info);
					}*/
                   /* else if (b.CallState== TAPI3Lib.CALL_STATE.CS_CONNECTED)
                    {
						MessageBox.Show("Connected");
						this.Alert("Connected", Form_Alert.enmType.Info);
					}*/
                   
                     if (b.CallState == TAPI3Lib.CALL_STATE.CS_OFFERING)
                    {
						this.Alert(cnumber, Form_Alert.enmType.Success);
						textBox1.Text = cnumber;
						MessageBox.Show(cnumber);
						this.Alert("Success", Form_Alert.enmType.Success);	
					}
					else if (b.CallState == TAPI3Lib.CALL_STATE.CS_DISCONNECTED)
					{
						this.Alert("Disconnected", Form_Alert.enmType.Warning);
						MessageBox.Show("Disconnected");

					}

					/*else if (b.CallState== TAPI3Lib.CALL_STATE.CS_IDLE)
                    {
						MessageBox.Show("Call is Created");
						this.Alert("Call is created!", Form_Alert.enmType.Info);
					}*/

					break;
			}
		}
	}
	
	//			case TAPI3Lib.TAPI_EVENT.TE_CALLSTATE:
	//				TAPI3Lib.ITCallStateEvent a = (TAPI3Lib.ITCallStateEvent)eobj;
	//				TAPI3Lib.ITCallInfo b = a.Call;
	//				switch (b.CallState)
	//				{
	//					case TAPI3Lib.CALL_STATE.CS_INPROGRESS:
	//						addtolist("dialing");
	//						break;
	//					case TAPI3Lib.CALL_STATE.CS_CONNECTED:
	//						addtolist("Connected");
	//						break;
	//					case TAPI3Lib.CALL_STATE.CS_DISCONNECTED:
	//						addtolist("Disconnected");
	//						break;
	//					case TAPI3Lib.CALL_STATE.CS_OFFERING:
	//						addtolist("A party wants to communicate with you!");
	//						break;
	//					case TAPI3Lib.CALL_STATE.CS_IDLE:
	//						addtolist("Call is created!");
	//						break;
	//				}
	//				break;
	//		}
	//	}
	//}
}
