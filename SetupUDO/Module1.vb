﻿Imports SAPbouiCOM.Framework
Imports SAPbobsCOM
Namespace UDOSetup

    Module Module1

        Public oCompany As SAPbobsCOM.Company
        Public sErrMsg As String
        Public lErrCode As Integer
        Public lRetCode As Integer
        Public sErrMsgV2 As String

        <STAThread()>
        Sub Main()
		
			'comment #1

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            Dim UDOTable As SAPbobsCOM.UserTablesMD
            Dim UDF_UDOTable As SAPbobsCOM.UserFieldsMD
            Dim sCookie As String
            Dim conStr As String
            Dim RetCode As Integer

            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String
            'Create the Documents object

            Dim vInvoice As SAPbobsCOM.Documents






            ' CUIDADO QUE LOS ODU DE 19 O MAS CARACTERES TIENE PROBLEMAS EN ALGUNAS VERSIONES.
            ' SE SUPONE RESUELTO EN 9.3 PL07.

            Try

                Dim oApp As Application

                oApp = New Application

                oCompany = New SAPbobsCOM.Company()
                sCookie = oCompany.GetContextCookie

                conStr = Application.SBO_Application.Company.GetConnectionContext(sCookie)
                oCompany.SetSboLoginContext(conStr)
                oCompany.Connect()

                Application.SBO_Application.StatusBar.SetText("Inicio proceso de creacion UDF", 2, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'Agregar UDF ItemCode a UDT @PPRECIOEXEP

                UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                UDF_UDOTable.TableName = "@PPRECIOEXEP"
                UDF_UDOTable.Name = "ItemDesc"
                UDF_UDOTable.Description = "Descripción Artículo"
                UDF_UDOTable.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                UDF_UDOTable.EditSize = 100


                RetCode = UDF_UDOTable.Add()

                oCompany.GetLastError(lRetCode, sErrMsg)

                If lRetCode Then
                    Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

                Exit Sub


            ' Creo UDO PERFILPRECIO

            UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            UDOTable.TableName = "PERFILPRECIO"
            UDOTable.TableDescription = "Perfil de Precios"
            UDOTable.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterData

            RetCode = UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDT: " & UDOTable.TableName & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Tabla UDT Creada" & " " & UDOTable.TableName, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDOTable)

            ' Creo UDO PPRECIOENTRIES

            UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            UDOTable.TableName = "PPRECIOENTRIES"
            UDOTable.TableDescription = "Descuentos"
            UDOTable.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterDataLines

            RetCode = UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDT: " & UDOTable.TableName & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Tabla UDT Creada" & " " & UDOTable.TableName, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDOTable)

            ' Creo UDO PPRECIOEXEP

            UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            UDOTable.TableName = "PPRECIOEXEP"
            UDOTable.TableDescription = "Item Excepciones"
            UDOTable.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterDataLines

            RetCode = UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDT: " & UDOTable.TableName & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Tabla UDT Creada" & " " & UDOTable.TableName, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDOTable)

            'Agregar UDF ListaDePrecios a UDT @PERFILPRECIO

            UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            UDF_UDOTable.TableName = "@PERFILPRECIO"
            UDF_UDOTable.Name = "ListaDePrecios"
            UDF_UDOTable.Description = "Lista de Precios"
            UDF_UDOTable.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
            UDF_UDOTable.EditSize = 5
            UDF_UDOTable.Mandatory = BoYesNoEnum.tYES

            RetCode = UDF_UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

            'Agregar UDF Rubro a UDT @PERFILPRECIO

            UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            UDF_UDOTable.TableName = "@PERFILPRECIO"
            UDF_UDOTable.Name = "Rubro"
            UDF_UDOTable.Description = "Rubro"
            UDF_UDOTable.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
            UDF_UDOTable.EditSize = 5
            UDF_UDOTable.Mandatory = BoYesNoEnum.tYES

            '// Adding the Field to the Table
            RetCode = UDF_UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

            'Agregar UDF Concepto a UDT @PPRECIOENTRIES

            UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)


            UDF_UDOTable.TableName = "@PPRECIOENTRIES"
            UDF_UDOTable.Name = "Concepto"
            UDF_UDOTable.Description = "Concepto"
            UDF_UDOTable.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            UDF_UDOTable.EditSize = 50
            UDF_UDOTable.Mandatory = BoYesNoEnum.tYES

            RetCode = UDF_UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

            'Agregar UDF Porcentaje a UDT @PPRECIOENTRIES

            UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            UDF_UDOTable.TableName = "@PPRECIOENTRIES"
            UDF_UDOTable.Name = "Porcentaje"
            UDF_UDOTable.Description = "Porcentaje"
            UDF_UDOTable.Type = BoFieldTypes.db_Float
            UDF_UDOTable.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage
            UDF_UDOTable.Mandatory = BoYesNoEnum.tYES

            RetCode = UDF_UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

            'Agregar UDF Computable a UDT @PPRECIOENTRIES

            UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            UDF_UDOTable.TableName = "@PPRECIOENTRIES"
            UDF_UDOTable.Name = "Computable"
            UDF_UDOTable.Description = "Computable"
            UDF_UDOTable.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            UDF_UDOTable.EditSize = 1
            UDF_UDOTable.Mandatory = BoYesNoEnum.tYES
            UDF_UDOTable.ValidValues.Value = "S"
            UDF_UDOTable.ValidValues.Description = "SI"
            UDF_UDOTable.ValidValues.Add()
            UDF_UDOTable.ValidValues.Value = "N"
            UDF_UDOTable.ValidValues.Description = "NO"
            UDF_UDOTable.ValidValues.Add()

            RetCode = UDF_UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

            'Agregar UDF ItemCode a UDT @PPRECIOEXEP

            UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            UDF_UDOTable.TableName = "@PPRECIOEXEP"
            UDF_UDOTable.Name = "ItemCode"
            UDF_UDOTable.Description = "Cod. Articulo"
            UDF_UDOTable.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            UDF_UDOTable.EditSize = 50
            UDF_UDOTable.Mandatory = BoYesNoEnum.tYES

            RetCode = UDF_UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

            'Agregar UDF Porcentaje a UDT @PPRECIOEXEP

            UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            UDF_UDOTable.TableName = "@PPRECIOEXEP"
            UDF_UDOTable.Name = "Porcentaje"
            UDF_UDOTable.Description = "Porcentaje"
            UDF_UDOTable.Type = BoFieldTypes.db_Float
            UDF_UDOTable.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage
            UDF_UDOTable.Mandatory = BoYesNoEnum.tYES

            RetCode = UDF_UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

            'Agregar UDF Computable a UDT @PPRECIOEXEP

            UDF_UDOTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            UDF_UDOTable.TableName = "@PPRECIOEXEP"
            UDF_UDOTable.Name = "Computable"
            UDF_UDOTable.Description = "Computable"
            UDF_UDOTable.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            UDF_UDOTable.EditSize = 1
            UDF_UDOTable.ValidValues.Value = "S"
            UDF_UDOTable.ValidValues.Description = "SI"
            UDF_UDOTable.ValidValues.Add()
            UDF_UDOTable.ValidValues.Value = "N"
            UDF_UDOTable.ValidValues.Description = "NO"
            UDF_UDOTable.ValidValues.Add()

            UDF_UDOTable.Mandatory = BoYesNoEnum.tYES

            RetCode = UDF_UDOTable.Add()

            oCompany.GetLastError(lRetCode, sErrMsg)

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDF: " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Campo UDF Creado" & " " & UDF_UDOTable.TableName & "." & UDF_UDOTable.Name, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(UDF_UDOTable)

            oUserObjectMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.FindColumns.ColumnAlias = "U_ListaDePrecios"
            oUserObjectMD.FindColumns.Add()
            oUserObjectMD.FindColumns.ColumnAlias = "U_Rubro"
            oUserObjectMD.FindColumns.Add()
            oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.TableName = "PERFILPRECIO"
            oUserObjectMD.Code = "PERFILPRECIO"
            oUserObjectMD.ChildTables.TableName = "PPRECIOENTRIES"
            oUserObjectMD.ChildTables.Add()
            oUserObjectMD.ChildTables.TableName = "PPRECIOEXEP"
            oUserObjectMD.ChildTables.Add()

            oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
            oUserObjectMD.Name = "Perfil de Precios"

            oUserObjectMD.Add()

            If lRetCode Then
                Application.SBO_Application.StatusBar.SetText("UDO: " & oUserObjectMD.Code & " " & sErrMsg, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Application.SBO_Application.StatusBar.SetText("Objeto UDO Creado: " & oUserObjectMD.Code, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            Catch ex As Exception
                Application.SBO_Application.StatusBar.SetText(ex.Message, 2, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try



        End Sub
    End Module
End Namespace

