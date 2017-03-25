Attribute VB_Name = "calls"
'Callback for run onAction
Sub call_openQEXfiles(control As IRibbonControl)
Call main.openQEXfiles
End Sub

'Callback for createMTO onAction
Sub call_MTOtemplate(control As IRibbonControl)
Call main.MTOtemplate
End Sub

'Callback for QTOConfig onAction
Sub call_QTOConfig(control As IRibbonControl)
Call STAGING.CONFIGsheet
End Sub

Sub call_SummaryQTO(control As IRibbonControl)
Call main.SummaryQTO
End Sub

'Callback for rules onAction
Sub call_runrules(control As IRibbonControl)
Call main.runrules
End Sub

'Callback for addrule onAction
Sub call_addrule(control As IRibbonControl)
createformula.Show
End Sub

'Callback for configrules onAction
Sub call_configrules(control As IRibbonControl)
Call STAGING.configrules
End Sub

'Callback for compare onAction
Sub call_costcodecomparison(control As IRibbonControl)
Call main.SummaryCOSTCODEComparison
End Sub

'Callback for PMreport onAction
Sub call_pmreport(control As IRibbonControl)
Call main.pmreport
End Sub

'Callback for tolconfig onAction
Sub call_configTol(control As IRibbonControl)
Call STAGING.CONFIGsheet
End Sub

'Callback for helpdoc onAction
Sub call_documentation(control As IRibbonControl)
    ThisWorkbook.FollowHyperlink ("https://docs.google.com/a/assemblesystems.com/document/d/1Dfkg171qgc7vHcLk9Oqw7x7j0EBuIY3qhqQvkMSBPb0/edit?usp=sharing")
End Sub
