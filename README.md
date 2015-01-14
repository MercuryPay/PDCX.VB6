PDCX.VB6 - MercuryPay Integration
========

Integrate to Mercury using DataCap's PDCX control.

#3 step process to integrate to PDCX.

##Step 1: Secure Device Initialization
  
This command should be performed during startup of the POS system with the optional PIN pad attached. 
It should not be performed prior to every transaction as it takes several seconds to complete with an attached PIN pad.
  
```
'' Create SecureDeviceInit XML, below example is for MagTek IPAD
 Dim request As String
 request = request + "<TStream>" + vbNewLine
 request = request + "   <Admin>" + vbNewLine
 request = request + "       <TranType>Setup</TranType>" + vbNewLine
 request = request + "       <TranCode>SecureDeviceInit</TranCode>" + vbNewLine
 request = request + "       <PadType>IPAD100</PadType>" + vbNewLine
 request = request + "       <SecureDevice>MTIPADHID</SecureDevice>" + vbNewLine
 request = request + "       <ComPort>0</ComPort>" + vbNewLine
 request = request + "   </Admin>" + vbNewLine
 request = request + "</TStream>"

'' Process the XML request
Dim response As String
response = Me.DsiPDCX1.ProcessTransaction(request, 0, "", "")
```
  
##Step 2: Process XML Transaction

Build XML transactions and process with PDCX object.

Below is a sample Credit Sale transaction.
```
'' MagTek IPAD Example XML
Dim request As String
request = request + "<?xml version=\"1.0\"?>" + vbNewLine
request = request + "<TStream>" + vbNewLine
request = request + "  <Transaction>" + vbNewLine
request = request + "    <MerchantID>019588466313922</MerchantID>" + vbNewLine
request = request + "    <LaneID>02</LaneID>" + vbNewLine
request = request + "    <TranType>Credit</TranType>" + vbNewLine
request = request + "    <TranCode>Sale</TranCode>" + vbNewLine
request = request + "    <InvoiceNo>10</InvoiceNo>" + vbNewLine
request = request + "    <RefNo>10</RefNo>" + vbNewLine    
request = request + "    <Frequency>OneTime</Frequency>" + vbNewLine
request = request + "    <RecordNo>RecordNumberRequested</RecordNo>" + vbNewLine
request = request + "    <PartialAuth>Allow</PartialAuth>" + vbNewLine
request = request + "    <Amount>" + vbNewLine
request = request + "      <Purchase>1.05</Purchase>" + vbNewLine
request = request + "    </Amount>" + vbNewLine
request = request + "    <SecureDevice>MTIPADHID</SecureDevice>" + vbNewLine
request = request + "    <ComPort>0</ComPort>" + vbNewLine
request = request + "    <Account>" + vbNewLine
request = request + "      <AcctNo>SecureDevice</AcctNo>" + vbNewLine
request = request + "    </Account>" + vbNewLine
request = request + "    <TerminalName>MPS Terminal</TerminalName>" + vbNewLine
request = request + "    <ShiftID>MPS Shift</ShiftID>" + vbNewLine
request = request + "    <OperatorID>MPS Operator</OperatorID>" + vbNewLine
request = request + "    <Memo>MPS PDCX Example v1.0</Memo>" + vbNewLine
request = request + "  </Transaction>" + vbNewLine
request = request + "</TStream>"

'' Process the XML request
Dim status As String
status = Me.DsiPDCX1..ServerIPConfig("x1.mercurydev.net;x2.mercurydev.net", 0);
Me.DsiPDCX1.SetConnectTimeout(5);
Me.DsiPDCX1.SetResponseTimeout(60);
Dim response As String
response = Me.DsiPDCX1.ProcessTransaction(request, 0, "", "")
```

##Step 3: Parse the XML Response

Approved transactions will have a CmdStatus equal to "Approved" or "Success".

```
Dim doc As New MSXML2.DOMDocument
doc.loadXML (response)

If doc.getElementsByTagName("CmdStatus").length > 0 Then

    If doc.getElementsByTagName("CmdStatus").Item(0).Text = "Success" _
        Or doc.getElementsByTagName("CmdStatus").Item(0).Text = "Approved" Then
        '' Approved logic
    Else
        '' Declined logic
    End If
    
Else
    '' Error logic
End If
```

###©2014 Mercury Payment Systems, LLC - all rights reserved.

Disclaimer:
This software and all specifications and documentation contained herein or provided to you hereunder (the "Software") are provided free of charge strictly on an "AS IS" basis. No representations or warranties are expressed or implied, including, but not limited to, warranties of suitability, quality, merchantability, or fitness for a particular purpose (irrespective of any course of dealing, custom or usage of trade), and all such warranties are expressly and specifically disclaimed. Mercury Payment Systems shall have no liability or responsibility to you nor any other person or entity with respect to any liability, loss, or damage, including lost profits whether foreseeable or not, or other obligation for any cause whatsoever, caused or alleged to be caused directly or indirectly by the Software. Use of the Software signifies agreement with this disclaimer notice.
