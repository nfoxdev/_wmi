*-----------------------------------------------------------------------------------
* Marco Plaza, 2022,2025
* https://github.com/nfoxdev/_wmi
* this program must be saved as _WMI.PRG
* v. 1.01
*-----------------------------------------------------------------------------------
*
* simple usage:
* _wmi( [wmiClassQ|wmi query] [, wmiNameSpace] )
*
* enter just the wmi class or specific wmi query; wmiNameSpace defaults to "CIMV2"
*
* oDisks = _wmi('Win32_diskDrive')  
* oMonitors = _wmi('Select * from Win32_PNPEntity where service = "monitor"') 
*
* 
* Test: save this program as _wmi.prg and do "testme in _wmi"
*
*------------------------------------------------------------------------------------
Lparameters wmiClassQ,wmiNameSpace

Local oerr,emessage,objwmiservice,oquery,owmi

wmiNamespace = Evl(m.wmiNamespace,'CIMV2')
emessage = ''

Try

  objwmiservice = Getobject("winmgmts:\\.\root\"+m.wmiNameSpace)

  If Getwordcount(m.wmiClassQ) = 1 or !lower(getwordnum(m.wmiClassQ,1)) == 'select'
    wmiquery = 'SELECT * FROM '+m.wmiClassQ
  Endif

  oquery = objwmiservice.execquery( m.wmiquery,,48)
  owmi = processobject( m.oquery )

Catch To oerr
  emessage = m.oerr.Message

Endtry

If !Empty(m.emessage)
  owmi = Null
  If Vartype(m.wmiClassQ) # 'C'
    wmiClassQ = 'invalid parameter type'
  Endif
  Error ' Invalid Query/wmiClassQ: '+m.wmiClassQ
Endif

Return m.owmi

*-------------------------------------------------
Procedure processobject( oquery )
*-------------------------------------------------
Local owmi,nitem,oitem

owmi = Createobject('empty')
AddProperty(owmi,'items(1)',.Null.)
nitem = 0


For Each oitem In m.oquery

  nitem = m.nitem + 1
  Dimension owmi.items(m.nitem)
  owmi.items(m.nitem) = Createobject('empty')
  setproperties( m.oitem, owmi.items(m.nitem) )

Endfor

AddProperty(owmi,'count',m.nitem)

Return m.owmi

*--------------------------------------------------------
Procedure setproperties( oitem , otarget  )
*--------------------------------------------------------

Local oerr,thisproperty,thisarray,nitem,thisitem,property,Item

For Each property In m.oitem.properties_

  Do Case
  Case Vartype( m.property.Value ) = 'O'
    thisproperty = Createobject('empty')
    setproperties(m.property.Value, m.thisproperty )
    AddProperty( otarget ,m.property.Name,m.thisproperty)

  Case m.property.isarray

    AddProperty( otarget ,property.Name+'(1)',.Null.)
    thisarray = 'otarget.'+m.property.Name

    nitem = 0

    If !Isnull(m.property.Value)

      For Each Item In m.property.Value

        nitem = m.nitem+1
        Dimension &thisarray(m.nitem)

        If Vartype( m.item) = 'O'
          thisitem = Createobject('empty')
          setproperties( m.item, m.thisitem )
          &thisarray(m.nitem) = m.thisitem
        Else
          &thisarray(m.nitem) = m.item
        Endif

      Endfor

    Endif

  Otherwise
    AddProperty( otarget ,m.property.Name,m.property.Value)
  Endcase

Endfor


*----------------------------------
Procedure testme
*----------------------------------
Public oinfo

oinfo = Create('empty')

Wait 'Running WMI Query....' Window Nowait At Wrows()/2,Wcols()/2

AddProperty( oinfo, "OperatingSystem"  	, _wmi('Win32_OperatingSystem') )
AddProperty( oinfo, "PhysicalMemory"  	, _wmi('Win32_PhysicalMemory') )
AddProperty( oinfo, "monitors"  		, _wmi('Win32_PNPEntity where service = "monitor"') )
AddProperty( oinfo, "diskdrive" 		, _wmi('Win32_diskDrive') )
AddProperty( oinfo, "startup" 			, _wmi('Win32_startupCommand'))
AddProperty( oinfo, "BaseBoard" 		, _wmi('Win32_baseBoard') )
AddProperty( oinfo, "netAdaptersConfig"	, _wmi('Win32_NetworkAdapterConfiguration') )

Messagebox( 'Please Inspect "oInfo" using debugger locals window ',0)
Debug
