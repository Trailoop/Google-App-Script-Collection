function getHtmlEmail () {
 var body = (<r><![CDATA[ 
     
     <!-- ***************************************************
********************************************************

HOW TO USE: Use these code examples as a guideline for formatting your HTML email. You may want to create your own template based on these snippets or just pick and choose the ones that fix your specific rendering issue(s). There are two main areas in the template: 1. The header (head) area of the document. You will find global styles, where indicated, to move inline. 2. The body section contains more specific fixes and guidance to use where needed in your design.

DO NOT COPY OVER COMMENTS AND INSTRUCTIONS WITH THE CODE to your message or risk spam box banishment :).

It is important to note that sometimes the styles in the header area should not be or don't need to be brought inline. Those instances will be marked accordingly in the comments.

********************************************************
**************************************************** -->

<!-- Using the xHTML doctype is a good practice when sending HTML email. While not the only doctype you can use, it seems to have the least inconsistencies. For more information on which one may work best for you, check out the resources below.

UPDATED: Now using xHTML strict based on the fact that gmail and hotmail uses it.  Find out more about that, and another great boilerplate, here: http://www.emailology.org/#1

More info/Reference on doctypes in email:
Campaign Monitor - http://www.campaignmonitor.com/blog/post/3317/correct-doctype-to-use-in-html-email/
Email on Acid - http://www.emailonacid.com/blog/details/C18/doctype_-_the_black_sheep_of_html_email_design
-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
	<title>Your Message Subject or Title</title>
	<style type="text/css">

		/***********
		Originally based on The MailChimp Reset from Fabio Carneiro, MailChimp User Experience Design
		More info and templates on Github: https://github.com/mailchimp/Email-Blueprints
		http://www.mailchimp.com &amp; http://www.fabio-carneiro.com

		INLINE: Yes.
		***********/
		/* Client-specific Styles */
		#outlook a {padding:0;} /* Force Outlook to provide a "view in browser" menu link. */
		body{width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%; margin:0; padding:0;}
		/* Prevent Webkit and Windows Mobile platforms from changing default font sizes, while not breaking desktop design. */
		.ExternalClass {width:100%;} /* Force Hotmail to display emails at full width */
		.ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height: 100%;} /* Force Hotmail to display normal line spacing.  More on that: http://www.emailonacid.com/forum/viewthread/43/ */
		#backgroundTable {margin:0; padding:0; width:100% !important; line-height: 100% !important;}
		/* End reset */

		/* Some sensible defaults for images
		1. "-ms-interpolation-mode: bicubic" works to help ie properly resize images in IE. (if you are resizing them using the width and height attributes)
		2. "border:none" removes border when linking images.
		3. Updated the common Gmail/Hotmail image display fix: Gmail and Hotmail unwantedly adds in an extra space below images when using non IE browsers. You may not always want all of your images to be block elements. Apply the "image_fix" class to any image you need to fix.

		Bring inline: Yes.
		*/
		img {outline:none; text-decoration:none; -ms-interpolation-mode: bicubic;}
		a img {border:none;}
		.image_fix {display:block;}

		/** Yahoo paragraph fix: removes the proper spacing or the paragraph (p) tag. To correct we set the top/bottom margin to 1em in the head of the document. Simple fix with little effect on other styling. NOTE: It is also common to use two breaks instead of the paragraph tag but I think this way is cleaner and more semantic. NOTE: This example recommends 1em. More info on setting web defaults: http://www.w3.org/TR/CSS21/sample.html or http://meiert.com/en/blog/20070922/user-agent-style-sheets/

		Bring inline: Yes.
		**/
		p {margin: 1em 0;}

		/** Hotmail header color reset: Hotmail replaces your header color styles with a green color on H2, H3, H4, H5, and H6 tags. In this example, the color is reset to black for a non-linked header, blue for a linked header, red for an active header (limited support), and purple for a visited header (limited support).  Replace with your choice of color. The !important is really what is overriding Hotmail's styling. Hotmail also sets the H1 and H2 tags to the same size.

		Bring inline: Yes.
		**/
		h1, h2, h3, h4, h5, h6 {color: black !important;}

		h1 a, h2 a, h3 a, h4 a, h5 a, h6 a {color: blue !important;}

		h1 a:active, h2 a:active,  h3 a:active, h4 a:active, h5 a:active, h6 a:active {
			color: red !important; /* Preferably not the same color as the normal header link color.  There is limited support for psuedo classes in email clients, this was added just for good measure. */
		 }

		h1 a:visited, h2 a:visited,  h3 a:visited, h4 a:visited, h5 a:visited, h6 a:visited {
			color: purple !important; /* Preferably not the same color as the normal header link color. There is limited support for psuedo classes in email clients, this was added just for good measure. */
		}

		/** Outlook 07, 10 Padding issue: These "newer" versions of Outlook add some padding around table cells potentially throwing off your perfectly pixeled table.  The issue can cause added space and also throw off borders completely.  Use this fix in your header or inline to safely fix your table woes.

		More info: http://www.ianhoar.com/2008/04/29/outlook-2007-borders-and-1px-padding-on-table-cells/
		http://www.campaignmonitor.com/blog/post/3392/1px-borders-padding-on-table-cells-in-outlook-07/

		H/T @edmelly

		Bring inline: No.
		**/
		table td {border-collapse: collapse;}

		/** Remove spacing around Outlook 07, 10 tables

		More info : http://www.campaignmonitor.com/blog/post/3694/removing-spacing-from-around-tables-in-outlook-2007-and-2010/

		Bring inline: Yes
		**/
		table {
  padding: 0; }
  table tr {
    border-top: 1px solid #cccccc;
    background-color: white;
    margin: 0;
    padding: 0; }
    table tr:nth-child(2n) {
      background-color: #f8f8f8; }
    table tr th {
      font-weight: bold;
      border: 1px solid #cccccc;
      text-align: left;
      margin: 0;
      padding: 6px 13px; }
    table tr td {
      border: 1px solid #cccccc;
      text-align: left;
      margin: 0;
      padding: 6px 13px; }
    table tr th :first-child, table tr td :first-child {
      margin-top: 0; }
    table tr th :last-child, table tr td :last-child {
      margin-bottom: 0; }

		/* Styling your links has become much simpler with the new Yahoo.  In fact, it falls in line with the main credo of styling in email, bring your styles inline.  Your link colors will be uniform across clients when brought inline.

		Bring inline: Yes. */
		a {color: orange;}

		/* Or to go the gold star route...
		a:link { color: orange; }
		a:visited { color: blue; }
		a:hover { color: green; }
		*/

		/***************************************************
		****************************************************
		MOBILE TARGETING

		Use @media queries with care.  You should not bring these styles inline -- so it's recommended to apply them AFTER you bring the other stlying inline.

		Note: test carefully with Yahoo.
		Note 2: Don't bring anything below this line inline.
		****************************************************
		***************************************************/

		/* NOTE: To properly use @media queries and play nice with yahoo mail, use attribute selectors in place of class, id declarations.
		table[class=classname]
		Read more: http://www.campaignmonitor.com/blog/post/3457/media-query-issues-in-yahoo-mail-mobile-email/
		*/
		@media only screen and (max-device-width: 480px) {

			/* A nice and clean way to target phone numbers you want clickable and avoid a mobile phone from linking other numbers that look like, but are not phone numbers.  Use these two blocks of code to "unstyle" any numbers that may be linked.  The second block gives you a class to apply with a span tag to the numbers you would like linked and styled.

			Inspired by Campaign Monitor's article on using phone numbers in email: http://www.campaignmonitor.com/blog/post/3571/using-phone-numbers-in-html-email/.

			Step 1 (Step 2: line 224)
			*/
			a[href^="tel"], a[href^="sms"] {
						text-decoration: none;
						color: black; /* or whatever your want */
						pointer-events: none;
						cursor: default;
					}

			.mobile_link a[href^="tel"], .mobile_link a[href^="sms"] {
						text-decoration: default;
						color: orange !important; /* or whatever your want */
						pointer-events: auto;
						cursor: default;
					}
		}

		/* More Specific Targeting */

		@media only screen and (min-device-width: 768px) and (max-device-width: 1024px) {
			/* You guessed it, ipad (tablets, smaller screens, etc) */

			/* Step 1a: Repeating for the iPad */
			a[href^="tel"], a[href^="sms"] {
						text-decoration: none;
						color: blue; /* or whatever your want */
						pointer-events: none;
						cursor: default;
					}

			.mobile_link a[href^="tel"], .mobile_link a[href^="sms"] {
						text-decoration: default;
						color: orange !important;
						pointer-events: auto;
						cursor: default;
					}
		}

		@media only screen and (-webkit-min-device-pixel-ratio: 2) {
			/* Put your iPhone 4g styles in here */
		}

		/* Following Android targeting from:
		http://developer.android.com/guide/webapps/targeting.html
		http://pugetworks.com/2011/04/css-media-queries-for-targeting-different-mobile-devices/  */
		@media only screen and (-webkit-device-pixel-ratio:.75){
			/* Put CSS for low density (ldpi) Android layouts in here */
		}
		@media only screen and (-webkit-device-pixel-ratio:1){
			/* Put CSS for medium density (mdpi) Android layouts in here */
		}
		@media only screen and (-webkit-device-pixel-ratio:1.5){
			/* Put CSS for high density (hdpi) Android layouts in here */
		}
		/* end Android targeting */
	</style>

	<!-- Targeting Windows Mobile -->
	<!--[if IEMobile 7]>
	<style type="text/css">

	</style>
	<![endif]-->

	<!-- ***********************************************
	****************************************************
	END MOBILE TARGETING
	****************************************************
	************************************************ -->

	<!--[if gte mso 9]>
	<style>
		/* Target Outlook 2007 and 2010 */
	</style>
	<![endif]-->
</head>
<body>
	<!-- Wrapper/Container Table: Use a wrapper table to control the width and the background color consistently of your email. Use this approach instead of setting attributes on the body tag. -->
	<table cellpadding="0" cellspacing="0" border="0" id="backgroundTable" style="border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; ">
	<tr>
		<td>
          <p><strong>TO</strong>: IDW, CDW, and SCDW Users<br/>
<strong>FM</strong>: MicroStrategy Support </p>

<p>The MicroStrategy support team will be offerings training sessions for all data warehouses (DW) including Item <abbr title="Item Data Warehouse">(IDW)</abbr>, Supply Chain <abbr title="Supply Chain Data Warehouse">(SCDW)</abbr>, and Customer <abbr title="Customer Data Warehouse">(CDW)</abbr> in Carlisle this June. </p>

<table>
<caption id="microstrategysupporttrainingofferings">MicroStrategy Support Training Offerings</caption>
<colgroup>
<col style="text-align:left;"/>
<col style="text-align:center;"/>
<col style="text-align:center;"/>
<col style="text-align:center;"/>
</colgroup>

<thead>
<tr>
	<th style="text-align:left;">Courses</th>
	<th style="text-align:center;">June 11</th>
	<th style="text-align:center;">June 12</th>
	<th style="text-align:center;">June 13</th>
</tr>
</thead>

<tbody>
<tr>
	<td style="text-align:left;" colspan="4">Introduction Sessions</td>
</tr>
<tr>
	<td style="text-align:left;">MicroStrategy/IDW</td>
	<td style="text-align:center;">8:30am - 2:00pm</td>
	<td style="text-align:center;"></td>
	<td style="text-align:center;">8:30am - 2:00pm</td>
</tr>
<tr>
	<td style="text-align:left;">SCDW</td>
	<td style="text-align:center;">2:30pm - 4:30pm</td>
	<td style="text-align:center;">9:00am - 11:00am</td>
	<td style="text-align:center;"></td>
</tr>
<tr>
	<td style="text-align:left;">CDW</td>
	<td style="text-align:center;"></td>
	<td style="text-align:center;">3:00pm - 4:00pm</td>
	<td style="text-align:center;"></td>
</tr>
<tr>
	<td style="text-align:left;" colspan="4">Other Sessions</td>
</tr>
<tr>
	<td style="text-align:left;">Intermediate MicroStrategy/IDW</td>
	<td style="text-align:center;"></td>
	<td style="text-align:center;">1:00pm - 3:00pm</td>
	<td style="text-align:center;"></td>
</tr>
<tr>
	<td style="text-align:left;">One On One</td>
	<td style="text-align:center;"></td>
	<td style="text-align:center;">11:00am - 12:00pm</td>
	<td style="text-align:center;">2:00pm &#8211;3:00pm</td>
</tr>
<tr>
	<td style="text-align:left;">Where</td>
	<td style="text-align:center;" colspan="3">All courses are being held in Carlisle Conf&#8211;363 PC Training</td>
</tr>
</tbody>
</table>
<h3>Courses Sign up Information</h3>

<p>Sign up for one of these offerings through the training classroom based registration <a href="https://ws1.aholdusa.com/jginet/cfappldata/trainingCourse/">here</a>. All classes are under <em>Ahold USA</em>. This registration requires access rights to the internal intranet. If you are having trouble signing up or do not have access rights contact us for help <a href="&#109;&#97;&#105;&#x6c;&#116;&#x6f;&#58;&#115;&#99;&#111;&#x74;&#116;&#46;&#108;&#97;&#119;&#114;&#x65;&#x6e;&#x63;&#101;&#64;&#97;&#x68;&#111;&#108;&#x64;&#117;&#x73;&#97;&#x2e;&#x63;&#x6f;&#109;&#x3f;&#115;&#117;&#98;&#106;&#x65;&#x63;&#x74;&#61;&#x52;&#x65;&#x67;&#x69;&#x73;&#116;&#114;&#x61;&#x74;&#105;&#111;&#110;&amp;&#110;&#x62;&#x73;&#112;&#x3b;&#65;&#x73;&#x73;&#x69;&#x73;&#x74;&#x61;&#x6e;&#x63;&#x65;&amp;&#x6e;&#x62;&#x73;&#112;&#59;&#78;&#x65;&#101;&#x64;&#101;&#x64;&amp;&#x62;&#x6f;&#x64;&#x79;&#61;&#83;&#99;&#x6f;&#x74;&#x74;">&#x68;&#101;&#x72;&#x65;</a>.</p>

<p>One On One sessions can be signed up for <a href="https://docs.google.com/a/ahold.com/spreadsheet/viewform?formkey=dDJxY1BTQ2lBOW85dFJDakMwTHlwZkE6MQ">here</a>.</p>

<h3>Course Descriptions</h3>

<h4>Introduction To MicroStrategy and Item Data Warehouse (IDW)</h4>

<p>Discover how to extract decision driving metrics and start impressing your peers and superiors.
You will learn the foundational concepts needed to design Microstrategy reports through the use of the Item Data Warehouse (IDW). IDW provides provides a number of report types that focus on providing point-of-sale (POS) information. The class provides ample opportunities to practice designing reports and organizing the data prior to extract. Attendees are encouraged to bring report ideas to get direct benefits from attending; and it makes for an interesting class. </p>

<p>approximate length: 4 hours 20 minutes | 1 hour lunch break | two 10 minute breaks provided. </p>

<h4>Introduction To Supply Chain Data Warehouse (SCDW)</h4>

<p>Need to verify shipment details, see the inventory in our warehouses, or mine data from purchase orders? Learn about the types of information available to you in the Supply Chain Data Warehouse (SCDW). This course focuses on reviewing the concepts needed for designing reports from the Supply Chain data warehouse. Attendees are encouraged to bring report ideas to get direct benefits from attending; and it makes for an interesting class.</p>

<p>approximate length: 2 hours | No breaks. </p>

<h4>Introduction To Consumer Data Warehouse (CDW)</h4>

<p>Learn about our customers by leveraging the loyalty card information collected in this data warehouse. Learn about our customer segmentation groups, trip missions categorization, and more. Attendees will focus on designing some basic CDW report types including Product Reports, Market Basket Reports, and Coupon Reports. Attendees are encouraged to bring report ideas to get direct benefits from attending; and it makes for an interesting class.</p>

<p>approximate length: 1 hour | No breaks.</p>

<h4>Advance/Intermediate MicroStrategy and IDW</h4>

<p>Looking for more knowledge of MicroStrategy features and other Item Data report types not covered in the beginner class? Learn about MicroStrategy features that allow you multiple item lists, try ranking and comparison report types. This class provides attendees time to build and optimize reports with onsite support. Attendees are encourages to bring reports they want to optimize or build and submit them <a href="&#109;&#97;&#105;&#x6c;&#x74;&#111;&#58;&#x73;&#99;&#x6f;&#x74;&#116;&#46;&#108;&#97;&#x77;&#x72;&#101;&#110;&#x63;&#x65;&#x40;&#97;&#x68;&#x6f;&#108;&#100;&#x75;&#x73;&#97;&#46;&#99;&#111;&#109;">&#x68;&#101;&#114;&#x65;</a> before class.</p>

<h3>Knowledge Base</h3>

<p>Learn more about MicroStrategy on the <a href="http://mstrweb.aholdusa.com/Website/Main_Page.htm">support site</a>.
Direct Links to <a href="http://mstrweb.aholdusa.com/Website/IDW%20Website/IDW_Home_Page.htm">IDW</a>, <a href="http://mstrweb.aholdusa.com/Website/CDW%20Website/CDW_Rpt_Team/CDW_Home_Page.htm">CDW</a>, and <a href="http://mstrweb.aholdusa.com/Website/SCDW%20Website/SCDW_Home_Page.htm">SCDW</a> resources.</p>

<h3>Need Assistance</h3>

<p>Having trouble finding the data you are looking for on MicroStrategy? Help is just an email away, <a href="&#109;&#x61;&#105;&#x6c;&#116;&#x6f;&#x3a;&#x41;&#x55;&#x53;&#65;&#x2e;&#x4d;&#x69;&#99;&#114;&#x6f;&#83;&#116;&#114;&#x61;&#116;&#x65;&#103;&#x79;&#x53;&#117;&#x70;&#x70;&#111;&#x72;&#x74;&#46;&#x67;&#114;&#x6f;&#x75;&#112;&#x40;&#x61;&#x68;&#111;&#108;&#x64;&#x2e;&#x63;&#x6f;&#109;">&#65;&#x55;&#x53;&#x41;&#x2e;&#x4d;&#x69;&#99;&#114;&#111;&#83;&#x74;&#x72;&#x61;&#116;&#101;&#x67;&#x79;&#83;&#x75;&#112;&#x70;&#111;&#x72;&#x74;&#x2e;&#103;&#x72;&#x6f;&#x75;&#x70;&#x40;&#97;&#104;&#x6f;&#108;&#x64;&#46;&#99;&#111;&#x6d;</a></p>
		</td>
	</tr>
	</table>
	<!-- End of wrapper table -->
</body>
</html> ]]></r>).toString();

return body;
}

function sendEmail() {
  var body = getHtmlEmail();
  var email = {
    "receipients": "scott.lawrence@aholdusa.com",
    "subject": "Upcoming Beginner Courses at QCP",
    "name": "MicroStrategy Support Group",
    "noReply": true
  };
  
  MailApp.sendEmail(email.receipients, email.subject, "", {name: email.name, htmlBody: body, noReply: email.noReply})
}
