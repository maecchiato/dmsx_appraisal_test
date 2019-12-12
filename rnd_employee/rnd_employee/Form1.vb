Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices
Imports System.Text

Public Class Form1
    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        searchDept()
    End Sub

    Private Sub connection()
        WTGCServer = "MIS07"
        WTGCPort = "1433"
        WTGCUsername = "sa"
        WTGCPassword = "intok"
        WTGCDatabase = "WTGC"

        WTGCConnectionString = "server=" & WTGCServer & "," & WTGCPort & ";uid=" & WTGCUsername & ";pwd=" & WTGCPassword & ";initial catalog=" & WTGCDatabase & ";"

    End Sub

    Public Sub searchDept()
        connection()

        lvRegList.BeginUpdate()
        lvRegList.Items.Clear()
        lvProbiList.BeginUpdate()
        lvProbiList.Items.Clear()
        Dim sQuery As String = "SELECT D.DName, A.AName, idMasterFile, unMasterFile, MF.unAccountUser, unPayrollSetupCustom, MF.unBusinessUnit, MF.unArea, MF.unDepartment, unBranch, unConfidentialityLevel, unBank, unChartTaxStatus, unJobTitle, MFEmployeeNumber, MFPayrollCategory, MFPayrollTerm, MFPayrollMode, MFCivilStatus, MFLaborGroup, MFEmploymentStatus, MFLastName, MFFirstName, MFMiddleName, MFSuffix, MFPhoto, MFBirthdate, MFAddress1, MFAddress2 , MFContactNumber1, MFContactNumber2, MFGender, MFDateHired, MFDateSeparated, MFDateCleared, MFSSS, MFPhilHealth, MFPagibig       , MFTaxID       , MFBankAccountNumber       , MFMonthlyRate       , MFDailyRate       , MFMonthlyAllowance, MFDailyAllowance, MFCola, MFTax, MF.[TimeStamp], MF.[Status]  FROM MasterFile AS MF 
                        Join Department as D on D.unDepartment = MF.unDepartment
                        Join Area as A on A.unArea = MF.unArea
                        WHERE MF.MFEmploymentStatus='Regular' AND D.DName like '%'+@deptsearch+'%' AND A.AName='Iloilo' Order By MFLastName Asc"
        Using WTGCConnection = New SqlConnection(WTGCConnectionString)
            WTGCConnection.Open()
            Using oCommand As New SqlCommand(sQuery, WTGCConnection)
                oCommand.Parameters.AddWithValue("@deptsearch", ComboBox1.Text)

                Dim oReader As SqlDataReader = oCommand.ExecuteReader
                While oReader.Read
                    Dim oItem As New ListViewItem
                    oItem.Text = oReader("MFEmployeeNumber")
                    oItem.SubItems.Add(oReader("MFLastName") + ", " + oReader("MFFirstName"))
                    lvRegList.Items.Add(oItem)
                End While
            End Using
        End Using

        Dim pQuery As String = "SELECT D.DName, A.AName, idMasterFile, unMasterFile, MF.unAccountUser, unPayrollSetupCustom, MF.unBusinessUnit, MF.unArea, MF.unDepartment, unBranch, unConfidentialityLevel, unBank, unChartTaxStatus, unJobTitle, MFEmployeeNumber, MFPayrollCategory, MFPayrollTerm, MFPayrollMode, MFCivilStatus, MFLaborGroup, MFEmploymentStatus, MFLastName, MFFirstName, MFMiddleName, MFSuffix, MFPhoto, MFBirthdate, MFAddress1, MFAddress2 , MFContactNumber1, MFContactNumber2, MFGender, MFDateHired, MFDateSeparated, MFDateCleared, MFSSS, MFPhilHealth, MFPagibig       , MFTaxID       , MFBankAccountNumber       , MFMonthlyRate       , MFDailyRate       , MFMonthlyAllowance, MFDailyAllowance, MFCola, MFTax, MF.[TimeStamp], MF.[Status]  FROM MasterFile AS MF 
                        Join Department as D on D.unDepartment = MF.unDepartment
                        Join Area as A on A.unArea = MF.unArea
                        WHERE MF.MFEmploymentStatus='Probationary' AND D.DName like '%'+@deptsearch+'%' AND A.AName='Iloilo' Order By MFLastName Asc"
        Using WTGCConnection = New SqlConnection(WTGCConnectionString)
            WTGCConnection.Open()
            Using oCommand As New SqlCommand(pQuery, WTGCConnection)
                oCommand.Parameters.AddWithValue("@deptsearch", ComboBox1.Text)
                Dim oReader As SqlDataReader = oCommand.ExecuteReader
                While oReader.Read
                    Dim oItem As New ListViewItem
                    oItem.Text = oReader("MFEmployeeNumber")
                    oItem.SubItems.Add(oReader("MFLastName") + ", " + oReader("MFFirstName"))
                    lvProbiList.Items.Add(oItem)
                End While
            End Using
        End Using
        lvProbiList.EndUpdate()
        lvRegList.EndUpdate()
    End Sub

    Public Sub searchEmployee()
        connection()
        lvRegList.BeginUpdate()
        lvRegList.Items.Clear()
        lvProbiList.BeginUpdate()
        lvProbiList.Items.Clear()
        Dim sQuery As String = "SELECT D.DName, A.AName, idMasterFile, unMasterFile, MF.unAccountUser, unPayrollSetupCustom, MF.unBusinessUnit, MF.unArea, MF.unDepartment, unBranch, unConfidentialityLevel, unBank, unChartTaxStatus, unJobTitle, MFEmployeeNumber, MFPayrollCategory, MFPayrollTerm, MFPayrollMode, MFCivilStatus, MFLaborGroup, MFEmploymentStatus, MFLastName, MFFirstName, MFMiddleName, MFSuffix, MFPhoto, MFBirthdate, MFAddress1, MFAddress2 , MFContactNumber1, MFContactNumber2, MFGender, MFDateHired, MFDateSeparated, MFDateCleared, MFSSS, MFPhilHealth, MFPagibig       , MFTaxID       , MFBankAccountNumber       , MFMonthlyRate       , MFDailyRate       , MFMonthlyAllowance, MFDailyAllowance, MFCola, MFTax, MF.[TimeStamp], MF.[Status]  FROM MasterFile AS MF 
                        Join Department as D on D.unDepartment = MF.unDepartment
                        Join Area as A on A.unArea = MF.unArea
                        WHERE MF.MFEmploymentStatus='Regular' AND D.DName like '%'+@deptsearch+'%' AND  MFLastName like '%'+@search+'%' OR MFFirstName like '%'+@search+'%' AND A.AName='Iloilo' Order By MFLastName Asc"
        Using WTGCConnection = New SqlConnection(WTGCConnectionString)
            WTGCConnection.Open()
            Using oCommand As New SqlCommand(sQuery, WTGCConnection)
                oCommand.Parameters.AddWithValue("@deptsearch", ComboBox1.Text)
                oCommand.Parameters.AddWithValue("@search", TextBox1.Text.Trim.Replace(" "c, "%"c))
                Dim oReader As SqlDataReader = oCommand.ExecuteReader
                While oReader.Read
                    Dim oItem As New ListViewItem
                    oItem.Text = oReader("MFEmployeeNumber")
                    oItem.SubItems.Add(oReader("MFLastName") + ", " + oReader("MFFirstName"))
                    lvRegList.Items.Add(oItem)
                End While
            End Using
        End Using

        Dim pQuery As String = "SELECT D.DName, A.AName, idMasterFile, unMasterFile, MF.unAccountUser, unPayrollSetupCustom, MF.unBusinessUnit, MF.unArea, MF.unDepartment, unBranch, unConfidentialityLevel, unBank, unChartTaxStatus, unJobTitle, MFEmployeeNumber, MFPayrollCategory, MFPayrollTerm, MFPayrollMode, MFCivilStatus, MFLaborGroup, MFEmploymentStatus, MFLastName, MFFirstName, MFMiddleName, MFSuffix, MFPhoto, MFBirthdate, MFAddress1, MFAddress2 , MFContactNumber1, MFContactNumber2, MFGender, MFDateHired, MFDateSeparated, MFDateCleared, MFSSS, MFPhilHealth, MFPagibig       , MFTaxID       , MFBankAccountNumber       , MFMonthlyRate       , MFDailyRate       , MFMonthlyAllowance, MFDailyAllowance, MFCola, MFTax, MF.[TimeStamp], MF.[Status]  FROM MasterFile AS MF 
                        Join Department as D on D.unDepartment = MF.unDepartment
                        Join Area as A on A.unArea = MF.unArea
                        WHERE MF.MFEmploymentStatus='Probationary' AND D.DName like '%'+@deptsearch+'%' AND MFLastName like '%'+@search+'%' OR MFFirstName like '%'+@search+'%' AND A.AName='Iloilo' Order By MFLastName Asc"
        Using WTGCConnection = New SqlConnection(WTGCConnectionString)
            WTGCConnection.Open()
            Using oCommand As New SqlCommand(pQuery, WTGCConnection)
                oCommand.Parameters.AddWithValue("@deptsearch", ComboBox1.Text)
                oCommand.Parameters.AddWithValue("@search", TextBox1.Text.Trim.Replace(" "c, "%"c))
                Dim oReader As SqlDataReader = oCommand.ExecuteReader
                While oReader.Read
                    Dim oItem As New ListViewItem
                    oItem.Text = oReader("MFEmployeeNumber")
                    oItem.SubItems.Add(oReader("MFLastName") + ", " + oReader("MFFirstName"))
                    lvProbiList.Items.Add(oItem)
                End While
            End Using
        End Using
        lvProbiList.EndUpdate()
        lvRegList.EndUpdate()
    End Sub

    Public Sub loadEmployee()
        connection()
        lvRegList.BeginUpdate()
        lvRegList.Items.Clear()
        lvProbiList.BeginUpdate()
        lvProbiList.Items.Clear()
        Dim sQuery As String = "SELECT D.Dname, idMasterFile, unMasterFile , MF.unAccountUser, unPayrollSetupCustom, MF.unBusinessUnit, MF.unArea, MF.unDepartment, unBranch, unConfidentialityLevel, unBank, unChartTaxStatus, unJobTitle, MFEmployeeNumber, MFPayrollCategory, MFPayrollTerm, MFPayrollMode, MFCivilStatus, MFLaborGroup, MFEmploymentStatus, MFLastName, MFFirstName, MFMiddleName, MFSuffix, MFPhoto, MFBirthdate, MFAddress1, MFAddress2, MFContactNumber1, MFContactNumber2, MFGender, MFDateHired, MFDateSeparated, MFDateCleared, MFSSS, MFPhilHealth, MFPagibig, MFTaxID, MFBankAccountNumber, MFMonthlyRate, MFDailyRate, MFMonthlyAllowance, MFDailyAllowance, MFCola, MFTax, MF.[TimeStamp], MF.[Status]  FROM  MasterFile as MF 
                                JOIN Department as D on D.unDepartment = MF.unDepartment 
                                Where MF.MFEmploymentStatus='Regular' AND D.DName like '%'+@deptsearch+'%' 
                                Order By MFLastName Asc"
        Using WTGCConnection = New SqlConnection(WTGCConnectionString)
            WTGCConnection.Open()
            Using oCommand As New SqlCommand(sQuery, WTGCConnection)
                oCommand.Parameters.AddWithValue("@deptsearch", ComboBox1.Text)
                Dim oReader As SqlDataReader = oCommand.ExecuteReader
                While oReader.Read
                    Dim oItem As New ListViewItem
                    oItem.Text = oReader("MFEmployeeNumber")
                    oItem.SubItems.Add(oReader("MFLastName") + ", " + oReader("MFFirstName"))
                    lvRegList.Items.Add(oItem)
                End While
            End Using
        End Using

        Dim pQuery As String = "SELECT D.Dname, idMasterFile, unMasterFile , MF.unAccountUser, unPayrollSetupCustom, MF.unBusinessUnit, MF.unArea, MF.unDepartment, unBranch, unConfidentialityLevel, unBank, unChartTaxStatus, unJobTitle, MFEmployeeNumber, MFPayrollCategory, MFPayrollTerm, MFPayrollMode, MFCivilStatus, MFLaborGroup, MFEmploymentStatus, MFLastName, MFFirstName, MFMiddleName, MFSuffix, MFPhoto, MFBirthdate, MFAddress1, MFAddress2, MFContactNumber1, MFContactNumber2, MFGender, MFDateHired, MFDateSeparated, MFDateCleared, MFSSS, MFPhilHealth, MFPagibig, MFTaxID, MFBankAccountNumber, MFMonthlyRate, MFDailyRate, MFMonthlyAllowance, MFDailyAllowance, MFCola, MFTax, MF.[TimeStamp], MF.[Status]  FROM  MasterFile as MF 
                                JOIN Department as D on D.unDepartment = MF.unDepartment 
                                Where MF.MFEmploymentStatus='Probationary' AND D.DName like '%'+@deptsearch+'%' 
                                Order By MFLastName Asc"
        Using WTGCConnection = New SqlConnection(WTGCConnectionString)
            WTGCConnection.Open()
            Using oCommand As New SqlCommand(pQuery, WTGCConnection)
                oCommand.Parameters.AddWithValue("@deptsearch", ComboBox1.Text)
                Dim oReader As SqlDataReader = oCommand.ExecuteReader
                While oReader.Read
                    Dim oItem As New ListViewItem
                    oItem.Text = oReader("MFEmployeeNumber")
                    oItem.SubItems.Add(oReader("MFLastName") + ", " + oReader("MFFirstName"))
                    lvProbiList.Items.Add(oItem)
                End While
            End Using
        End Using
        lvProbiList.EndUpdate()
        lvRegList.EndUpdate()
    End Sub

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        connection()
    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick
        If (TextBox1.Text = "Search") Then
            TextBox1.Clear()
            TextBox1.ForeColor = Color.Black
            searchDept()
        End If
    End Sub

    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If (TextBox1.Text = Nothing) Then
            TextBox1.Text = "Search"
            TextBox1.ForeColor = Color.LightGray
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'searchDept()
        searchEmployee()
    End Sub

    'Open the data from lvSRC
    Private Sub OpenNewEmpProfileTab(lv As ListView)
        For Each oTabSelect As TabPage In tcProfile.TabPages
            If oTabSelect.Text = lv.SelectedItems(0).Text Then
                tcProfile.SelectedTab = oTabSelect
                Return
            End If
        Next

        Dim oSED As New ucEmpData
        oSED.Dock = DockStyle.Fill
        'oSED.idOrderControlCommi = lvRegList.SelectedItems(0).Tag

        Dim oTab As New TabPage
        oTab.Text = lv.SelectedItems(0).Text
        oTab.Controls.Add(oSED)

        tcProfile.TabPages.Add(oTab)
        tcProfile.SelectedTab = oTab
    End Sub

    Private Sub lvRegList_DoubleClick(sender As Object, e As EventArgs) Handles lvRegList.DoubleClick
        OpenNewEmpProfileTab(lvRegList)
    End Sub

    Private Sub lvProbiList_DoubleClick(sender As Object, e As EventArgs) Handles lvProbiList.DoubleClick
        OpenNewEmpProfileTab(lvProbiList)
    End Sub
End Class
