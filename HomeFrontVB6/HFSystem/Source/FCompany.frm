VERSION 5.00
' TAG: Fixed dialog form for editing company information (name, address, phone, fax, email, tax number, etc.)
' CONVERT: Implement as a Blazor modal dialog component with _isVisible bool and overlay div pattern.
' CONVERT: Use private fields for each company property bound with @bind to input elements.
' CONVERT: Use EventCallback<CompanyData> OnSave parameter to return edited data to the parent component.
' CONVERT: Load country/province dropdowns via DbWrapperSqlServer QueryAsync on initialization.
' CONVERT: Handle Escape key via @onkeydown to close the dialog.
Begin VB.Form FCompany 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company"
   ClientHeight    =   3690
   ClientLeft      =   5055
   ClientTop       =   2100
   ClientWidth     =   5310
   Icon            =   "FCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   ' TAG: TextBox for company email address (index 11)
   ' CONVERT: Replace with <input type="text" @bind="Email" maxlength="50" /> in the Email field row.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   11
      Left            =   1695
      MaxLength       =   50
      TabIndex        =   7
      Text            =   " "
      Top             =   2145
      Width           =   3375
   End
   ' TAG: TextBox for county name (index 7)
   ' CONVERT: Replace with <input type="text" @bind="County" maxlength="50" /> in the County field row.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   6
      Text            =   " "
      Top             =   1755
      Width           =   3375
   End
   ' TAG: Multi-line TextBox for company address (index 2, up to 102 chars)
   ' CONVERT: Replace with <textarea rows="2" maxlength="102" @bind="Address" /> in the Address field row.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   2
      Left            =   1680
      MaxLength       =   102
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   570
      Width           =   3375
   End
   ' TAG: TextBox for city name (index 3)
   ' CONVERT: Replace with <input type="text" @bind="City" maxlength="50" /> in the City/Prov field row.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1005
      Width           =   2565
   End
   ' TAG: TextBox for postal code (index 5, max 7 chars)
   ' CONVERT: Replace with <input type="text" @bind="PostalCode" maxlength="7" style="max-width:140px;" />.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   230
      Index           =   5
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1260
      Width           =   1395
   End
   ' TAG: TextBox for phone number (index 8, IME disabled for numeric input)
   ' CONVERT: Replace with <input type="text" @bind="Phone" maxlength="30" style="max-width:180px;" />.
   ' CONVERT: Add @onblur handler to call FormatPhone() helper method.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   8
      Left            =   1695
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2400
      Width           =   1755
   End
   ' TAG: TextBox for company name (index 1)
   ' CONVERT: Replace with <input type="text" @bind="CompanyName" /> in the Company Name field row.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   1680
      TabIndex        =   0
      Text            =   " "
      Top             =   180
      Width           =   3375
   End
   ' TAG: TextBox for fax number (index 9, IME disabled for numeric input)
   ' CONVERT: Replace with <input type="text" @bind="Fax" maxlength="30" style="max-width:180px;" />.
   ' CONVERT: Add @onblur handler to call FormatPhone() helper method.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   9
      Left            =   1695
      MaxLength       =   30
      TabIndex        =   9
      Top             =   2655
      Width           =   1755
   End
   ' TAG: TextBox for tax number (index 10, IME disabled, max 20 chars)
   ' CONVERT: Replace with <input type="text" @bind="TaxNumber" maxlength="20" style="max-width:160px;" />.
   ' CONVERT: Add @onfocus handler for select-all behavior via JS interop.
   Begin VB.TextBox text1 
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   10
      Left            =   1695
      MaxLength       =   20
      TabIndex        =   10
      Top             =   3060
      Width           =   1575
   End
   ' TAG: Custom combo box for selecting the company's country
   ' CONVERT: Replace with <select @bind="Country" @bind:after="OnCountryChanged"> with options from CountryList.
   ' CONVERT: Load CountryList via DbWrapperSqlServer QueryAsync("SELECT DISTINCT country FROM countrycodes ORDER BY 1").
   ' CONVERT: On country change, clear and reload ProvinceList for the selected country.
   Begin HFSystem.VBCombo cboCompanyCountry 
      Height          =   240
      Left            =   1680
      TabIndex        =   5
      Top             =   1500
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   423
   End
   ' TAG: Custom combo box for selecting the company's province/state
   ' CONVERT: Replace with <select @bind="Province"> with options from ProvinceList.
   ' CONVERT: ProvinceList is dynamically loaded based on the selected Country value.
   Begin HFSystem.VBCombo cboCompanyProvince 
      Height          =   240
      Left            =   4260
      TabIndex        =   3
      Top             =   1005
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   423
   End
   ' TAG: Label for the Email field (right-aligned)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "Email".
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Email"
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   20
      Top             =   2190
      Width           =   1155
   End
   ' TAG: Label for the County field (right-aligned)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "County".
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "County"
      Height          =   195
      Index           =   0
      Left            =   405
      TabIndex        =   19
      Top             =   1800
      Width           =   1155
   End
   ' TAG: Label for the City/Province field (right-aligned, auto-sized)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "City/Prov".
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City/Prov"
      Height          =   195
      Left            =   900
      TabIndex        =   18
      Top             =   1020
      Width           =   660
   End
   ' TAG: Label for the Address field (right-aligned, auto-sized)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "Address".
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   990
      TabIndex        =   17
      Top             =   570
      Width           =   570
   End
   ' TAG: Label for the Postal Code field (right-aligned, auto-sized)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "Postal Code".
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Postal Code"
      Height          =   195
      Left            =   705
      TabIndex        =   16
      Top             =   1275
      Width           =   855
   End
   ' TAG: Label for the Company Name field (right-aligned, auto-sized)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "Company Name".
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Company Name"
      Height          =   195
      Index           =   51
      Left            =   435
      TabIndex        =   15
      Top             =   180
      Width           =   1125
   End
   ' TAG: Label for the Phone field (right-aligned)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "Phone".
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone"
      Height          =   195
      Index           =   47
      Left            =   420
      TabIndex        =   14
      Top             =   2400
      Width           =   1155
   End
   ' TAG: Label for the Fax field (right-aligned)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "Fax".
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax"
      Height          =   195
      Index           =   45
      Left            =   420
      TabIndex        =   13
      Top             =   2655
      Width           =   1155
   End
   ' TAG: Label for the Tax Number field (right-aligned)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "Tax Number".
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tax Number"
      Height          =   195
      Index           =   52
      Left            =   180
      TabIndex        =   12
      Top             =   3030
      Width           =   1395
   End
   ' TAG: Label for the Country field (right-aligned)
   ' CONVERT: Replace with <label> element inside a company-field-row div with text "Country".
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Country"
      Height          =   195
      Index           =   53
      Left            =   405
      TabIndex        =   11
      Top             =   1515
      Width           =   1155
   End
End
Attribute VB_Name = "FCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private field(11) As String

' TAG: Populates the form fields from a VSFlexGrid row, shows the form modally, then writes edited values back to the grid.
' CONVERT: Replace with a public async Task ShowForm(CompanyData data) method that populates bound properties.
' CONVERT: Use IsVisible bool + StateHasChanged() to show the modal dialog overlay.
' CONVERT: After dialog closes via OK, fire OnSave EventCallback<CompanyData> with the updated CompanyData object.
' CONVERT: Parse Address into Address1/Address2 using string.Split on newline when returning data.
' CONVERT: Remove VSFlexGrid dependency; parent component handles grid row read/write.
Public Sub ShowForm(Grid As VSFlexGrid)
    With Grid
        text1(1) = .TextMatrix(.Row, .ColIndex("Company"))
        text1(2) = .TextMatrix(.Row, .ColIndex("Address1")) & vbCrLf & _
                   .TextMatrix(.Row, .ColIndex("Address2"))
        text1(3) = .TextMatrix(.Row, .ColIndex("City"))
        text1(5) = .TextMatrix(.Row, .ColIndex("Postal"))
        Call SetComboBoxListIndex(cboCompanyCountry, .TextMatrix(.Row, .ColIndex("Country")))
        Call SetComboBoxListIndex(cboCompanyProvince, .TextMatrix(.Row, .ColIndex("Province")))
        text1(7) = .TextMatrix(.Row, .ColIndex("County"))
        text1(8) = .TextMatrix(.Row, .ColIndex("Phone"))
        text1(9) = .TextMatrix(.Row, .ColIndex("Fax"))
        text1(10) = .TextMatrix(.Row, .ColIndex("TaxNumber"))
        text1(11) = .TextMatrix(.Row, .ColIndex("Email"))
        Me.Show vbModal
        .TextMatrix(.Row, .ColIndex("Company")) = field(1)
        .TextMatrix(.Row, .ColIndex("Address1")) = Parse(field(2), 1, vbCrLf)
        .TextMatrix(.Row, .ColIndex("Address2")) = Parse(field(2), 2, vbCrLf)
        .TextMatrix(.Row, .ColIndex("City")) = field(3)
        .TextMatrix(.Row, .ColIndex("Province")) = field(4)
        .TextMatrix(.Row, .ColIndex("Postal")) = field(5)
        .TextMatrix(.Row, .ColIndex("Country")) = field(6)
        .TextMatrix(.Row, .ColIndex("County")) = field(7)
        .TextMatrix(.Row, .ColIndex("Phone")) = field(8)
        .TextMatrix(.Row, .ColIndex("Fax")) = field(9)
        .TextMatrix(.Row, .ColIndex("TaxNumber")) = field(10)
        .TextMatrix(.Row, .ColIndex("Email")) = field(11)
    End With
End Sub

' TAG: Delegates province combo change event to the click handler to update the province field value.
' CONVERT: Not needed in Blazor; @bind on the Province <select> handles value changes automatically.
Private Sub cboCompanyProvince_Change()
    Call cboCompanyProvince_Click
End Sub
' TAG: Updates the province field value when a province is selected from the combo box.
' CONVERT: Not needed in Blazor; @bind="Province" on the <select> element handles this automatically.
Private Sub cboCompanyProvince_Click()
    field(4) = cboCompanyProvince.Text
End Sub

' TAG: Closes the form when the Escape key is pressed.
' CONVERT: Handle via @onkeydown on the overlay div; if e.Key == "Escape", set IsVisible = false and invoke IsVisibleChanged.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
' TAG: Initializes the form by loading saved position/size and populating the country combo box from the database.
' CONVERT: Implement in OnInitializedAsync; call LoadCountries() which queries "SELECT DISTINCT country FROM countrycodes ORDER BY 1" via DbWrapperSqlServer.
' CONVERT: Form position/size persistence (IniGetForm) is not needed in Blazor web UI.
Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadComboBox(cboCompanyCountry, HFApp.Databases(dbHomefront), "select distinct country,'',0 from countrycodes order by 1")
End Sub
' TAG: Saves the form position/size to INI when the form is unloaded.
' CONVERT: Not needed in Blazor; form position persistence is not applicable to web modal dialogs.
Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub
' TAG: Syncs the field array with the text box value whenever a text box changes.
' CONVERT: Not needed in Blazor; use @bind on each input to directly bind to the corresponding property (CompanyName, Address, City, etc.).
Private Sub text1_Change(Index As Integer)
    field(Index) = text1(Index)
End Sub
' TAG: Selects all text in a text box when it receives focus.
' CONVERT: Implement via JS interop: call JSRuntime.InvokeVoidAsync to select text on @onfocus of each input.
Private Sub text1_GotFocus(Index As Integer)
    Call SelectAll(text1(Index))
End Sub

' TAG: Placeholder for country combo change event (currently commented out, delegates to click handler).
' CONVERT: Not needed in Blazor; @bind:after="OnCountryChanged" on the Country <select> handles this.
Private Sub cboCompanyCountry_Change()
 '   Call cboCompanyCountry_Click
End Sub
' TAG: Updates the country field and reloads the province combo box based on the selected country.
' CONVERT: Implement as async Task OnCountryChanged() method that clears Province, then calls LoadProvinces(Country).
' CONVERT: LoadProvinces should query "SELECT DISTINCT state FROM countrycodes WHERE country = @Country ORDER BY 1" via DbWrapperSqlServer.
' CONVERT: Bind the Country <select> with @bind:after="OnCountryChanged" to trigger province reload.
Private Sub cboCompanyCountry_Click()
    field(6) = cboCompanyCountry.Text
    cboCompanyProvince.Clear
    cboCompanyProvince.ListIndex = -1
    Call LoadComboBox(cboCompanyProvince, HFApp.Databases(dbHomefront), "select distinct state,'',0 from countrycodes where country=" & DbQuote(Str, cboCompanyCountry.Text) & "order by 1")
End Sub


' TAG: Formats phone and fax text box values on validation (when leaving the field).
' CONVERT: Implement as @onblur handlers on Phone and Fax inputs that call a FormatPhone() helper method.
' CONVERT: FormatPhone should strip non-digits and format as (xxx) xxx-xxxx for 10-digit numbers.
Private Sub text1_Validate(Index As Integer, Cancel As Boolean)
    If IsIn(Index, 8, 9) Then
        text1(Index).Text = FormatPhone(text1(Index).Text)
    End If
End Sub