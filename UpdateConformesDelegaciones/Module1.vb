Imports ComunesRedux

Module Module1

    Private baseDatos As New BaseDatos(BaseDatos.MODO_PRODUCCION)

    Sub Main()
        ''''Aplicación que se incluirá dentro del programador para actualizar los conformes de aquellas expediciones
        ''''Trasmitidas por EDI entre delegaciones, con esto evitamos tener que escanear la documentación dos veces.
        Dim exito As Integer = 0
        Console.WriteLine("Inicio proceso actualización conformes delegaciones")
        Dim strconsulta As String = "update ase  set ase.arcverrut=af.arcverrut, ase.arcverfec=getdate(), ase.arcvermat=1 " &
                    "from expedic4 PS " &
                    "inner join EXPEDIC4 PF ON ps.HolCod=pf.HolCod and " &
                    "pf.ExpAlbOrd=cast(ps.SecCod as varchar(2)) + '-' + cast(ps.ExpCtrCod as varchar(3)) + '-' + cast(ps.ExpCod as varchar(6)) " &
                    "inner join arcver ASe ON ase.HolCod=ps.HolCod and ase.ArcVerCtr=ps.ExpCtrCod and ase.ArcVerSec=ps.SecCod and " &
                    "ase.ArcVerNiv = 3 And ase.ArcVerSub = 1 And ase.ArcVerEmp = 0 And ((ase.ArcVerCod = 1 and ase.arcversec<>15) or (ase.ArcVerCod = 8 and ase.arcversec in (15,23,24))) And ase.ArcVerRef = ps.ExpCod " &
                    "inner join arcver AF ON af.HolCod=pf.HolCod and af.ArcVerCtr=pf.ExpCtrCod and af.ArcVerSec=pf.SecCod and " &
                    "af.ArcVerNiv = 3 And af.ArcVerSub = 1 And af.ArcVerEmp = 0 And af.ArcVerCod = 1 And af.ArcVerRef = pf.ExpCod " &
                    "where ps.ExpDatEtd>DATEADD(day,-40,getdate()) and ps.holcod=0 and not ase.ArcVerRut like '\\%' and af.ArcVerRut like '\\%'"
        exito = instSQL.SentenciaSQL(strconsulta, baseDatos.sConexionTRANS)
        If exito = 0 Then
            strconsulta = "update ase  set ase.arcverrut=af.arcverrut, ase.arcverfec=getdate(), ase.arcvermat=1 " &
                   "from expedic4 PS " &
                   "inner join EXPEDIC4 PF ON ps.HolCod=pf.HolCod and " &
                   "pf.ExpAlbOrd=cast(ps.SecCod as varchar(2)) + '-' + cast(ps.ExpCtrCod as varchar(3)) + '-' + cast(ps.ExpCod as varchar(6)) " &
                   "inner join arcver ASe ON ase.HolCod=ps.HolCod and ase.ArcVerCtr=ps.ExpCtrCod and ase.ArcVerSec=ps.SecCod and " &
                   "ase.ArcVerNiv = 3 And ase.ArcVerSub = 1 And ase.ArcVerEmp = 0 And ((ase.ArcVerCod = 1 and ase.arcversec<>15) or (ase.ArcVerCod = 8 and ase.arcversec in (15,23,24))) And ase.ArcVerRef = ps.ExpCod " &
                   "inner join arcver AF ON af.HolCod=pf.HolCod and af.ArcVerCtr=pf.ExpCtrCod and af.ArcVerSec=pf.SecCod and " &
                   "af.ArcVerNiv = 3 And af.ArcVerSub = 1 And af.ArcVerEmp = 0 And af.ArcVerCls=91 And af.ArcVerRef = pf.ExpCod " &
                   "where ps.ExpDatEtd>DATEADD(day,-40,getdate()) and ps.holcod=0 and not ase.ArcVerRut like '\\%' and af.ArcVerRut like '\\%'"
            exito = instSQL.SentenciaSQL(strconsulta, baseDatos.sConexionTRANS)
        End If

        strconsulta = "update ase  set ase.arcverrut=af.arcverrut, ase.arcverfec=getdate(), ase.arcvermat=1 " &
                    "from expedic4 PS " &
                    "inner join EXPEDIC4 PF ON ps.HolCod=pf.HolCod and " &
                    "substring(pf.ExpAlbOrd,1,18) = cast(ps.SecCod as varchar(2)) + '3029' + RIGHT('000' + cast(ps.ExpCtrCod as varchar(3)),3) + RIGHT('000000000' + cast(ps.ExpCod as varchar(9)),9) " &
                    "inner join arcver ASe ON ase.HolCod=ps.HolCod and ase.ArcVerCtr=ps.ExpCtrCod and ase.ArcVerSec=ps.SecCod and " &
                    "ase.ArcVerNiv = 3 And ase.ArcVerSub = 1 And ase.ArcVerEmp = 0 And ((ase.ArcVerCod = 1 and ase.arcversec<>15) or (ase.ArcVerCod = 8 and ase.arcversec in (15,23,24))) And ase.ArcVerRef = ps.ExpCod " &
                    "inner join arcver AF ON af.HolCod=pf.HolCod and af.ArcVerCtr=pf.ExpCtrCod and af.ArcVerSec=pf.SecCod and " &
                    "af.ArcVerNiv = 3 And af.ArcVerSub = 1 And af.ArcVerEmp = 0 And af.ArcVerCod = 1 And af.ArcVerRef = pf.ExpCod " &
                    "where ps.ExpDatEtd>DATEADD(day,-40,getdate()) and ps.holcod=0 and not ase.ArcVerRut like '\\%' and af.ArcVerRut like '\\%'"
        exito = instSQL.SentenciaSQL(strconsulta, baseDatos.sConexionTRANS)
        If exito = 0 Then
            strconsulta = "update ase  set ase.arcverrut=af.arcverrut, ase.arcverfec=getdate(), ase.arcvermat=1 " &
                   "from expedic4 PS " &
                   "inner join EXPEDIC4 PF ON ps.HolCod=pf.HolCod and " &
                   "substring(pf.ExpAlbOrd,1,18) = cast(ps.SecCod as varchar(2)) + '3029' + RIGHT('000' + cast(ps.ExpCtrCod as varchar(3)),3) + RIGHT('000000000' + cast(ps.ExpCod as varchar(9)),9) " &
                   "inner join arcver ASe ON ase.HolCod=ps.HolCod and ase.ArcVerCtr=ps.ExpCtrCod and ase.ArcVerSec=ps.SecCod and " &
                   "ase.ArcVerNiv = 3 And ase.ArcVerSub = 1 And ase.ArcVerEmp = 0 And ((ase.ArcVerCod = 1 and ase.arcversec<>15) or (ase.ArcVerCod = 8 and ase.arcversec in (15,23,24))) And ase.ArcVerRef = ps.ExpCod " &
                   "inner join arcver AF ON af.HolCod=pf.HolCod and af.ArcVerCtr=pf.ExpCtrCod and af.ArcVerSec=pf.SecCod and " &
                   "af.ArcVerNiv = 3 And af.ArcVerSub = 1 And af.ArcVerEmp = 0 And af.ArcVerCls=91 And af.ArcVerRef = pf.ExpCod " &
                   "where ps.ExpDatEtd>DATEADD(day,-40,getdate()) and ps.holcod=0 and not ase.ArcVerRut like '\\%' and af.ArcVerRut like '\\%'"
            exito = instSQL.SentenciaSQL(strconsulta, baseDatos.sConexionTRANS)
        End If

        Console.WriteLine("Fin Proceso actualización")
    End Sub

End Module
