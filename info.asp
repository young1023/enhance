<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="imagetoolbar" content="no">
<TITLE>Jaguar Pen Limited</TITLE>
<LINK title=style1 
href="images/basic.css" type=text/css rel=stylesheet><!--BEGIN menuItemData.asp -->
<SCRIPT language=JavaScript 
src="images/js_dynMenuFunctions.js">
<!-- Functions used to create the dynamic HTML drop-down menus -->
</SCRIPT>

<!--#include file="menu_bar.inc"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<font color="#000000">
<SCRIPT language=javascript>
<!--
//function displayAWindow(url, width, height){
//var Win=window.open(url,"displayAWindow",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=no');
//}

	var gstrBrowserType = null
	if( document.layers )   gstrBrowserType="Netscape";
	else if( document.all )	gstrBrowserType="IE";
	else if( document.getElementById ) gstrBrowserType = "NS6";
	gstrMenuColor = "white"
	gstrMenuColorOn = "#E5DED9"
	gstrBorderColor = "#E41F1F"
	var strAlignment = "left"

	if( gstrBrowserType == "Netscape" )
	{
		style = "gNetMenuText"

		for (var intCnt=0; intCnt<garyMenuNames.length; intCnt++) {
			if( intCnt == (garyMenuNames.length - 1) ){
				//All menus are left-aligned except the last menu
				strAlignment = "left"
				gstrBorderColor = "#4F2403"
				style = "gNetMenuLastText"
			}
			//Pass the menu name, the menu left position Netscape, the Netscape width, and the alignment of text in the menu
			createNetscapeMenu(garyMenuNames[intCnt][0], garyMenuNames[intCnt][1], garyMenuNames[intCnt][3], strAlignment, style)
		}
	}
	else if( gstrBrowserType == "IE" )
	{
		if( "" == "AOL")
		{
			style = "gAOLMenuItemStyle"
		}
		else
		{
			style = "gIEMenuItemStyle"
		}

		for (var intCnt=0; intCnt<garyMenuNames.length; intCnt++) {
			if( intCnt == (garyMenuNames.length - 1) ){
				//All menus are left-aligned except the last menu
				strAlignment = "left"
				gstrBorderColor = "#4F2403"
				style = "gIEMenuLastItemStyle"
			}

			//Pass the menu name, the menu left position for IE, the IE width, and the alignment of text in the menu
			createIEMenu(garyMenuNames[intCnt][0], garyMenuNames[intCnt][2], garyMenuNames[intCnt][4], strAlignment, style)
		}
	}
	else if( gstrBrowserType == "NS6" )
	{
		style = "gIEMenuItemStyle"

		for (var intCnt=0; intCnt<garyMenuNames.length; intCnt++) {
			if( intCnt == (garyMenuNames.length - 1) ){
				//All menus are left-aligned except the last menu
				strAlignment = "left"
				gstrBorderColor = "#4F2403"
				style = "gIEMenuLastItemStyle"
			}

			//Pass the menu name, the menu left position for IE, the IE width, and the alignment of text in the menu
			createNSsixMenu(garyMenuNames[intCnt][0], garyMenuNames[intCnt][5], garyMenuNames[intCnt][6], strAlignment, style)
		}
	}


function openWin(arg_url, arg_name, arg_width, arg_height, arg_scroll, arg_resize) {
    var win = window.open(arg_url,arg_name,'scrollbars=' + arg_scroll + ',resizable=' + arg_resize + ',width=' + arg_width +  ',height=' + arg_height);
  }
//-->
</SCRIPT>
<!--#include file="menu_nav.inc"-->

<SCRIPT language=JavaScript>
<!--
var dcs_imgarray = new Array;
var dcs_ptr = 0;
var dCurrent = new Date();
var DCS=new Object();
var WT=new Object();
var DCSext=new Object();
var saveReqArgs = "";

//var dcsADDR = "64.15.240.148";
var dcsADDR = "metric.organic.com";
var dcsID = "";

if (dcsID == ""){
	var TagPath = dcsADDR;
} else {
	var TagPath = dcsADDR+"/"+dcsID;
}

function dcs_var(){
	WT.tz = dCurrent.getTimezoneOffset();
	WT.ul = navigator.appName=="Netscape" ? navigator.language : navigator.userLanguage;
	WT.cd = screen.colorDepth;
	WT.sr = screen.width+"x"+screen.height;
	WT.jo = navigator.javaEnabled() ? "Yes" : "No";
	WT.ti   = document.title;
	DCS.dcsdat = dCurrent.getTime();
	if ((window.document.referrer != "") && (window.document.referrer != "-")){
		if (!(navigator.appName == "Microsoft Internet Explorer" && parseInt(navigator.appVersion) < 4) ){
			DCS.dcsref = window.document.referrer;
		}
	}
    	if ( !window.forcePathname) {
	    DCS.dcsuri = window.location.pathname;
	    //alert ('window.location.pathname = ' + window.location.pathname);
	} else {
	    DCS.dcsuri = forcePathname;
	    //alert ('forcePathname = ' + forcePathname);
	}
	DCS.dcsqry = window.location.search;
	DCS.dcssip = window.location.hostname;
}

function A(N,V){
	return "&"+N+"="+escape(V);
}

function dcs_createImage(dcs_src)
{
	if (document.images){
		dcs_imgarray[dcs_ptr] = new Image;
		dcs_imgarray[dcs_ptr].src = dcs_src;
		dcs_ptr++;
	}
}

function dcs_TAG(TagImage){
	var P ="http"+(window.location.protocol.indexOf('https:')==0?'s':'')+"://"+TagImage+"/dcs.gif?";
	for (N in DCS){P+=A( N, DCS[N]);}
	for (N in WT){P+=A( "WT."+N, WT[N]);}
	for (N in DCSext){P+=A( N, DCSext[N]);}
	saveReqArgs = P;
	dcs_createImage(P);
}

dcs_var();
dcs_TAG(TagPath);
//-->
</SCRIPT>
      <NOSCRIPT><IMG height=1 src="images/njs.gif" width=1 
      border=0 name=DCSIMG> </NOSCRIPT>        
</font>       
      <TABLE cellSpacing=0 cellPadding=0 width=758 height=400 border=0>
        <TBODY>
        <TR>
          <TD>

<!-- content is inserted here ---->
        
<table width="98%" border="0" cellspacing="5" cellpadding="2">
  <tr> 
    <td><b><font color="#000000">Your Privacy Assured</font></b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">Your privacy is of great concern to <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>.  
      We use the information you provide to process your order and to ensure that  
      your experience on our site is as enjoyable and efficient as possible.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">The Information We Gather and How We Use It</font></b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">Our system gathers information  
      about the areas you visit on our Internet store. We do not use any of the  
      navigational data that you - as an individual - provide while browsing or  
      shopping on the Internet Store; nor do we share this data with anyone outside  
      our company. We do use navigational information in the aggregate to understand  
      how our visitors shop in our store so we can make <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>  
      better.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">In order to fulfill your order and provide you with highest level of customer  
      service, we must ask you for personal information. <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>  
      collects your name, shipping and billing addresses, telephone number, e-mail  
      address information in order to deliver your order promptly and to the correct  
      address. We use your e-mail address to send you a confirmation of your order.  
      We will also send an e-mail when your package is shipped and include the  
      shipper's tracking number. You may also choose to register with us. Registration  
      provides advantages like faster check out; ability to store shipping addresses;  
      advanced notification on upcoming promotions; and exclusive invitations  
      to private sales and events. By registering, you can also update your personal  
      information through your <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>  
      account. If you place an order, registering allows you to track your order  
      status, but is not required to complete the order process. <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>  
      does not disclose, share, sell, or trade any personal or individual information  
      that it collects to third parties, other than in the course of completing  
      a transaction, providing you with information that you have requested, or  
      in response to an investigation involving legal matters. While we may provide  
      aggregate information regarding site usage to third parties, such as statistics  
      on site traffic and sales, you can rest assured that we will only work with  
      external companies who will respect our customer's privacy.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Marketing Policy</font></b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">You have a choice regarding how the information you have provided may  
      be used to make special offers to you. You can direct to remove your name,  
      address, and e-mail address from mailing lists, we provide to select <b>Subscribe</b>  
      or <b>Unsubscribe</b> at the member profile page :<a href="http://asp.jaguarpen.com/add_member1.asp">  
      http://asp.jaguarpen.com/add_member1.asp</a></font></td>
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">If you do not wish to receive information on special Jaguar Pen offers  
      and promotions, please e-mail us at <a href="mailto:storemaster@jaguarpen.com">storemaster@jaguarpen.com</a>  
      You may also mail these requests to: Jaguar Pen (HK) Ltd Attn: Marketing  
      Department, 1/F, Air Goal Cargo Bldg., 330 Kwun Tong Road, Kowloon, Hong 
      Kong. Once we receive  
      your request, we will promptly remove you from our marketing list. If your  
      interests or needs change in the future, please let us know.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Safety and Privacy of Minors</font></b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">Jaguar Pen does not target children nor does <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>  
      expect to attract children to its site. Because of this, our privacy measures  
      have not been specifically developed for their protection on the Internet.  
      We do not submit information  
      to <a href="http://www.jaguarpen.com">www.jaguarpen.com</a> without the  
      express consent and participation of a parent or legal guardian.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Links</font></b></td>
  </tr>
  <tr> 
    <td><font color="#000000">Occasionally, this site may contain links to other sites. <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>  
      is not responsible for the privacy practices or the content of other web  
      sites. We suggest you check the privacy statement on each site that you  
      visit.</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Contacting the Website</font></b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">If you have any questions or concerns about the <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>  
      privacy policy, contact us at <a href="mailto:storemaster@jaguarpen.com">storemaster@jaguarpen.com</a>,  
      or write to us at: 1/F, Air Goal Cargo Bldg., 330 Kwun Tong Road, Kowloon, 
      Hong Kong.</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Secure Transactions</font></b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">Statistically, it's safer to shop online with your credit card than to  
      use it in a restaurant or department store. At <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>,  
      protecting your personal information is extremely important to us. To ensure  
      that you never have to worry about credit card safety when you shop in our  
      store, we use World Pay Online Payment Gateway <a href="http://www.worldpay.com">www.worldpay.com</a>  
      and its 128 bit Secure Sockets Layer (SSL) security level server, which  
      is the industry standard and among the best software available today for  
      secure commerce transactions. When the shopper fills in the payment form  
      and clicks the submit button, their details are not sent straight away.  
      What actually happens is that a secure link is set up between the shopper's  
      browser and WorldPay and an encryption code is requested and received, which  
      then wraps the order and transaction details before leaving the shoppers  
      premises. To see a fuller example of this in action, please <a href="http://www.worlddpay.com/uk/products%20services/how%20it%20works.shtml">click  
      here</a></font></td>
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">Once Worldpay receives your credit card information, it is immediately  
      encrypted and stored in their secure data center. This step prevents anyone  
      except the credit card authorization company from viewing it. After you've  
      given them your credit card information, it is never re-displayed on our  
      site. To ensure your protection, even our customer service associates have  
      access to only show the transaction ID, customer account ID etc, no credit  
      card number shown.We will send you an immediate email verifying receipt  
      of your order with your order confirmation number. Please keep this number  
      handy if you need to contact us with any questions about your order.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Exchanges and Returns</font></b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">At <a href="http://www.jaguarpen.com">www.jaguarpen.com</a>, we'd like  
      to help you resolve a purchase problem as quickly and easily as possible.  
      Provided you meet the terms and conditions listed below, you may make an  
      exchange or return. Please note that <b>not</b> all items are covered by  
      our returns policy. Before you make your purchase read the complete returns  
      policy information.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">General Policies</font></b></td> 
  </tr>
  <tr> 
    <td> <font color="#000000">&#149; Shipping charges are non-refundable.  
      Normally 25% of the sample charge amount.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&#149; All exchanges and returns require  
      a Return Authorization No. It will be sent via email to you. Exchanges or  
      return cannot be processed without it.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&#149; Request exchange or return merchandise  
      should be within 30 days of the original invoice date. Thereafter, all sales  
      are final.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&#149; The returned package must be received  
      within 10 business days of the Return Authorization issued date.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&#149; When returning product, we strongly  
      recommend the use of a carrier that can track packages. You also assume  
      responsibility for insuring the returned item.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&#149; All items must be returned in the  
      original packaging and have all accessories, blank warranty cards, and owner's  
      manuals.</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">&#149; Please note all returns conditions  
      listed below, as applicable to your return.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&#149; We do not charge any restocking fees.</font>  
    </td>
  </tr>
  <tr> 
    <td><font color="#000000">&#149; Over weight items (10 kgs, or greater)  
      need to be coordinated through our <a href="mailto:storemaster@jaguarpen.com">storemaster@jaguarpen.com</a></font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Exceptions to our Returns Policy</font></b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">We cannot accept returns for some products, please views the restrictions  
      as below :</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">At times, <a href="http://www.jaguarpen.com">www.jaguarpen.com</a> will  
      offer specially priced items that have been discontinued by the manufacturer.  
      These items, clearly marked as discontinued, are not eligible for return  
      or exchange.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">Returns on clearance items are accepted within 30 days of the invoice  
      date if the product is damaged or deemed defective. An RA must be obtained  
      in order for a clearance return to be accepted and processed. We cannot  
      exchange clearance products, as the factory has either discontinued these  
      items or we no longer have them in stock.</font> </td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Requesting A Product Return</font> </b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">You can choose any of these convenient methods:</font> </td> 
  </tr>
  <tr> 
    <td><font color="#000000">1. Send your request via email to <a href="mailto:storemaster@jaguarpen.com">storemaster@jaguarpen.com</a>.  
      Please include your order number.</font> </td> 
  </tr>
  <tr> 
    <td><font color="#000000">2. You may contact us at (852) 2798 6263 from 9AM to 6PM, Monday through  
      Friday and Saturday, 9AM to 1PM Hong Kong time (GMT+8 hours)</font> </td> 
  </tr>
  <tr> 
    <td><font color="#000000">3. Fax us at (852) 2751 6659</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">4. Or write to : Jaguar Pen (HK) Ltd 1/F, Air Goal 
      Cargo Bldg.,330 Kwun Tong Road, Kowloon, Hong Kong</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">After Receiving Your Return Authorization Number</font>  
      </b> </td>
  </tr>
  <tr> 
    <td><font color="#000000">1. Write your Return Authorization number legibly on the return label.</font>  
    </td>
  </tr>
  <tr> 
    <td><font color="#000000">2. Return the entire package to:<b>JAGUAR PEN (HK) LTD</b>. G/F, On Tai  
      Bldg., 47-49 Ting On St., Kowloon, Hong Kong.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Refund</font></b></td>
  </tr>
  <tr> 
    <td><font color="#000000">Refunds will be given at the discretion of the Company Management.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">Credit will be issued within two business days of receiving your return.  
      All refunds for credit will be issued to the credit card account that appears  
      on the original invoice. Shipping and handling fees are non-refundable.  
      Please note that your financial institution may take up to 10-15 additional  
      days from the date we issue the credit to post it to your actual account.  
      Questions regarding this should be directed to your financial institution.</font>  
    </td>
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Product Replacement</font> </b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">When exchanging an item, you will be sent a replacement item and a new  
      tracking number once Jaguar Pen receives, approves and processes the returned  
      item.</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Sample Order Acceptance</font></b></td> 
  </tr>
  <tr> 
    <td height="10"><font color="#000000">Jaguar Pen appreciates your business. Our business is structured  
      to supply the writing instruments product needs of individuals, businesses,  
      institutions or agencies for their &quot;commercial sample&quot; or &quot;personal&quot;  
      use only. Jaguar Pen accept orders from dealers, exporters, wholesalers,  
      or other customers who intend to resell the products that we offer for sale.</font>  
    </td>
  </tr>
  <tr> 
    <td><font color="#000000"><a href="http://www.jaguarpen.com">www.jaguarpen.com</a> reserves the  
      right to, at any time, after receipt of any order, from any customer, and  
      without prior notice, supply less than the quantity ordered of any item  
      or to cancel an order.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Payment</font></b></td>
  </tr>
  <tr> 
    <td><font color="#000000"><a href="http://www.jaguarpen.com">www.jaguarpen.com</a> accepts <b>credit  
      cards</b> as formal methods of payment. We are currently accept <b>Visa</b>,  
      <b>MasterCard</b>, <b>Delta, and</b> <b>Visa Electron</b>.  
      No surcharge for using your credit card with us:</font> </td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><b><font color="#000000">Please Note:</font> </b></td> 
  </tr>
  <tr> 
    <td><font color="#000000">From time to time, products can be constrained or discontinued. We will  
      update you on a weekly basis with the status of your backorder. If after  
      30 days we are unable to fulfill your backorder, we will cancel your backorder,  
      credit your credit card, and notify you accordingly. You may cancel your  
      backorder at any time, whereupon we will credit your credit card account.</font></td> 
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">&nbsp;</font></td>
  </tr>
  <tr>
    <td>
     <div align="center"><font size="2" color="#000000">Copyright&copy; 2001 Winson  
  Development Company All right reserved.</font> </div> 
    </td>
  </tr>
</table>

<!-- content is ended here ---->

</TD>









<!--#include file="menu_bottom.inc"-->
</BODY>
