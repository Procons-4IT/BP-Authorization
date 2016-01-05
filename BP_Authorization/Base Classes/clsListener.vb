Public Class clsListener
    Inherits Object

    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter

#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property

#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        'oApplication.SBO_Application.SetFilter(Filters)
    End Sub

    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters()

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            objFilter.Add(frm_GRPO) 'GRPO
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_SaleReturn) 'Return
            objFilter.AddEx(frm_ARCreditMemo) 'AR Credit Memo
            objFilter.AddEx(frm_ORDR) 'Order
            objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            objFilter.AddEx(frm_Quotation) ' Quotation
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment
            objFilter.AddEx(frm_ARReverseInvoice) 'Reverse Invoice
            objFilter.AddEx(frm_ChoosefromList) ' Choose From List          
           
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            objFilter.AddEx(frm_ChoosefromList) ' Choose From List          

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            objFilter.Add(frm_GRPO) 'GRPO
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_SaleReturn) 'Return
            objFilter.AddEx(frm_ARCreditMemo) 'AR Credit Memo
            objFilter.AddEx(frm_ORDR) 'Order
            objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            objFilter.AddEx(frm_Quotation) ' Quotation
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment
            objFilter.AddEx(frm_ARReverseInvoice) 'Reverse Invoice
            objFilter.AddEx(frm_ChoosefromList) ' Choose From List            

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED)
            objFilter.Add(frm_GRPO) 'GRPO
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_SaleReturn) 'Return
            objFilter.AddEx(frm_ARCreditMemo) 'AR Credit Memo
            objFilter.AddEx(frm_ORDR) 'Order
            objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            objFilter.AddEx(frm_Quotation) ' Quotation
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment
            objFilter.AddEx(frm_ARReverseInvoice) 'Reverse Invoice

            SetFilter(objFilters)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

#End Region

#End Region

#Region "Data Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.FormTypeEx
                
            End Select
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Menu Event"

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                End Select
            Else
                Select Case pVal.MenuUID
                    Case mnu_CLOSE
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                    Case mnu_DUPLICATE
                        Dim oForm As SAPbouiCOM.Form
                        oForm = oApplication.SBO_Application.Forms.Item(_FormUID)
                    Case mnu_DELETE_ROW
                        Dim oForm As SAPbouiCOM.Form
                        oForm = oApplication.SBO_Application.Forms.Item(_FormUID)
                End Select
            End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub

#End Region

#Region "Item Event"

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID
            If pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                Select Case pVal.FormTypeEx
                    Case frm_ChoosefromList
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsChooseFromList
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_GRPO, frm_APCreditMemo, frm_APDownPaymentInvoice, frm_APDownPaymentReq, frm_APInvoice, frm_APReserverInvoice, frm_PurchaseOrder, frm_PurchaseQuotation
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsGRPO
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Quotation
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsQuotation
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ORDR
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsOrder
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Delivery
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDelivery
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_INVOICES
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsInvoice
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_INVOICESPAYMENT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsInvoicePayment
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SaleReturn
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsReturn
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ARCreditMemo
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsARCreditMemo
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ARReverseInvoice
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsReverseInvoice
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_IncomingPayment
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsIncomingPayment
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_OutPayment
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsOutGoingPayment
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                End Select
            End If
            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If

                Dim oForm As SAPbouiCOM.Form
                If pVal.FormTypeEx = frm_ChoosefromList Then
                    _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
                End If
                If (pVal.FormTypeEx = frm_Quotation.ToString() Or pVal.FormTypeEx = frm_ORDR.ToString() Or pVal.FormTypeEx = frm_Delivery.ToString Or pVal.FormTypeEx = frm_SaleReturn.ToString Or pVal.FormTypeEx = frm_INVOICES.ToString Or pVal.FormTypeEx = frm_INVOICESPAYMENT.ToString Or pVal.FormTypeEx = frm_ARCreditMemo.ToString Or pVal.FormTypeEx = frm_ARDownPaymentReq.ToString Or pVal.FormTypeEx = frm_ARDownPaymentInvoice.ToString) And (pVal.BeforeAction = True And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED)) Then 'Validate Accounts for Production Costing
                    If 1 = 1 Then
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        If (pVal.ItemUID = "4" Or pVal.ItemUID = "54") And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim oComboBox As SAPbouiCOM.ComboBox
                            oComboBox = oForm.Items.Item("2001").Specific
                            Try
                                If oComboBox.Selected.Description = "" Then
                                    oApplication.Utilities.Message("Branch is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                oApplication.Utilities.Message("Branch is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                            If oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) <> "" Then

                                If oApplication.Utilities.validateBP(oApplication.Utilities.getEdittextvalue(oForm, "4"), oApplication.Utilities.getEdittextvalue(oForm, "54"), oComboBox.Selected.Value) = False Then
                                    '  oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                    '  oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                Else
                                    Exit Sub
                                End If
                            End If

                            If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                Dim objChooseForm As SAPbouiCOM.Form
                                Dim objChoose As New clsChooseFromList
                                clsChooseFromList.ItemUID = "4" '.ItemUID
                                clsChooseFromList.CardNameUID = "54"
                                clsChooseFromList.SourceFormUID = FormUID
                                clsChooseFromList.SourceLabel = 0
                                clsChooseFromList.CFLChoice = "Customer"
                                clsChooseFromList.choice = "Customer"
                                '  Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oForm.Items.Item("2001").Specific
                                Try
                                    clsChooseFromList.Branch = oComboBox.Selected.Value
                                Catch ex As Exception
                                    clsChooseFromList.Branch = ""
                                End Try

                                clsChooseFromList.sourceColumID = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value
                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(objChooseForm)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                oComboBox = oForm.Items.Item("2001").Specific
                                Try
                                    If oComboBox.Selected.Description = "" Then
                                        oApplication.Utilities.Message("Branch is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                Catch ex As Exception
                                    oApplication.Utilities.Message("Branch is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End Try
                                If oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) <> "" Then
                                    If oApplication.Utilities.validateBP(oApplication.Utilities.getEdittextvalue(oForm, "4"), oApplication.Utilities.getEdittextvalue(oForm, "54"), oComboBox.Selected.Value) = False Then
                                        '   oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                    Else
                                        Exit Sub

                                    End If
                                End If
                                Dim objChooseForm As SAPbouiCOM.Form
                                Dim objChoose As New clsChooseFromList
                                clsChooseFromList.ItemUID = pVal.ItemUID
                                clsChooseFromList.ItemUID = "4" '.ItemUID
                                clsChooseFromList.CardNameUID = "54"
                                clsChooseFromList.SourceFormUID = FormUID
                                clsChooseFromList.SourceLabel = 0
                                clsChooseFromList.CFLChoice = "Customer"
                                clsChooseFromList.choice = "Customer"
                                ' Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oForm.Items.Item("2001").Specific
                                Try
                                    clsChooseFromList.Branch = oComboBox.Selected.Value
                                Catch ex As Exception
                                    clsChooseFromList.Branch = ""
                                End Try

                                clsChooseFromList.sourceColumID = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value
                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(objChooseForm)
                                BubbleEvent = False
                                Exit Sub
                                End If
                        End If
                    End If
                ElseIf (pVal.FormTypeEx = frm_PurchaseQuotation.ToString() Or pVal.FormTypeEx = frm_PurchaseOrder.ToString() Or pVal.FormTypeEx = frm_GRPO.ToString Or pVal.FormTypeEx = frm_GoodsReturn.ToString Or pVal.FormTypeEx = frm_APDownPaymentReq.ToString Or pVal.FormTypeEx = frm_APDownPaymentInvoice.ToString Or pVal.FormTypeEx = frm_APInvoice.ToString Or pVal.FormTypeEx = frm_APCreditMemo.ToString) And (pVal.BeforeAction = True And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED)) Then 'Validate Accounts for Production Costing
                    If 1 = 1 Then
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        If (pVal.ItemUID = "4" Or pVal.ItemUID = "54") And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Dim oComboBox As SAPbouiCOM.ComboBox
                            oComboBox = oForm.Items.Item("2001").Specific
                            Try
                                If oComboBox.Selected.Description = "" Then
                                    oApplication.Utilities.Message("Branch is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                oApplication.Utilities.Message("Branch is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub

                            End Try
                            If oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) <> "" Then
                                If oApplication.Utilities.validateBP(oApplication.Utilities.getEdittextvalue(oForm, "4"), oApplication.Utilities.getEdittextvalue(oForm, "54"), oComboBox.Selected.Value) = False Then
                                    ' oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                    ' oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                Else
                                    Exit Sub
                                End If
                            End If

                            If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                Dim objChooseForm As SAPbouiCOM.Form
                                Dim objChoose As New clsChooseFromList
                                clsChooseFromList.ItemUID = "4" '.ItemUID
                                clsChooseFromList.CardNameUID = "54"
                                clsChooseFromList.SourceFormUID = FormUID
                                clsChooseFromList.SourceLabel = 0
                                clsChooseFromList.CFLChoice = "Vendor"
                                clsChooseFromList.choice = "Vendor"
                                ' Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oForm.Items.Item("2001").Specific
                                Try
                                    clsChooseFromList.Branch = oComboBox.Selected.Value
                                Catch ex As Exception
                                    clsChooseFromList.Branch = ""
                                End Try
                                clsChooseFromList.sourceColumID = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value
                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(objChooseForm)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Dim objChooseForm As SAPbouiCOM.Form
                                Dim objChoose As New clsChooseFromList
                                clsChooseFromList.ItemUID = "4" '.ItemUID
                                clsChooseFromList.CardNameUID = "54"
                                clsChooseFromList.SourceFormUID = FormUID
                                clsChooseFromList.SourceLabel = 0
                                clsChooseFromList.CFLChoice = "Vendor"
                                clsChooseFromList.choice = "Vendor"
                                ' Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oForm.Items.Item("2001").Specific
                                Try
                                    clsChooseFromList.Branch = oComboBox.Selected.Value
                                Catch ex As Exception
                                    clsChooseFromList.Branch = ""
                                End Try
                                clsChooseFromList.sourceColumID = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value
                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(objChooseForm)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    ElseIf (pVal.FormTypeEx = frm_IncomingPayment.ToString() Or pVal.FormTypeEx = frm_OutPayment.ToString()) And (pVal.BeforeAction = True And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED)) Then 'Validate Accounts for Production Costing
                        If 1 = 1 Then
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            If (pVal.ItemUID = "5" Or pVal.ItemUID = "32") And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim strValue As String = String.Empty
                                If (CType(oForm.Items.Item("56").Specific, SAPbouiCOM.OptionBtn).Selected) Then
                                    strValue = "Customer"

                                End If
                                If (CType(oForm.Items.Item("57").Specific, SAPbouiCOM.OptionBtn).Selected) Then
                                strValue = "Vendor"
                                End If

                                Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oForm.Items.Item("1320002037").Specific
                                Try
                                    If oComboBox.Selected.Description = "" Then
                                        oApplication.Utilities.Message("Branch is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                Catch ex As Exception
                                    oApplication.Utilities.Message("Branch is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                            End Try
                            If oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) <> "" Then

                                If oApplication.Utilities.validateBP(oApplication.Utilities.getEdittextvalue(oForm, "5"), oApplication.Utilities.getEdittextvalue(oForm, "32"), oComboBox.Selected.Value) = False Then
                                    ' oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                Else
                                    Exit Sub
                                End If
                            End If

                            If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                Dim objChooseForm As SAPbouiCOM.Form
                                Dim objChoose As New clsChooseFromList
                                clsChooseFromList.ItemUID = "5" '.ItemUID
                                clsChooseFromList.CardNameUID = "32"
                                clsChooseFromList.SourceFormUID = FormUID
                                clsChooseFromList.SourceLabel = 0
                                clsChooseFromList.CFLChoice = strValue
                                clsChooseFromList.choice = strValue
                                '  Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oForm.Items.Item("1320002037").Specific
                                Try
                                    clsChooseFromList.Branch = oComboBox.Selected.Value
                                Catch ex As Exception
                                    clsChooseFromList.Branch = ""
                                End Try
                                clsChooseFromList.sourceColumID = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value
                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(objChooseForm)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Dim objChooseForm As SAPbouiCOM.Form
                                Dim objChoose As New clsChooseFromList
                                clsChooseFromList.ItemUID = "5" '.ItemUID
                                clsChooseFromList.CardNameUID = "32"
                                clsChooseFromList.SourceFormUID = FormUID
                                clsChooseFromList.SourceLabel = 0
                                clsChooseFromList.CFLChoice = strValue
                                clsChooseFromList.choice = strValue
                                '  Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oForm.Items.Item("1320002037").Specific
                                Try
                                    clsChooseFromList.Branch = oComboBox.Selected.Value
                                Catch ex As Exception
                                    clsChooseFromList.Branch = ""
                                End Try

                                clsChooseFromList.sourceColumID = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value
                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, "")
                                oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(objChooseForm)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    Else
                        Try
                            _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
                        Catch ex As Exception

                        End Try

                    End If
            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If
                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If
            End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

#Region "Right Click Event"

    Private Sub _SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.RightClickEvent
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_ORDR Then
                oMenuObject = New clsOrder
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Application Event"

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub

#End Region

#Region "Close Application"

    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

#Region "Set Application"

    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub

#End Region

#Region "Finalize"

    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region
   
End Class
