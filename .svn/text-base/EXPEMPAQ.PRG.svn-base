***********************************************
*                                             *
* Autor: Francisco Alcayde Rosas.             *
* Empresa: Vianney, S.A. de C.V.              *
* Fecha: 28 de Abril de 1998.                 *
* Discusion: Procedimiento que genera el      *
*            paquin list de un embarque de    *
*            exportacion y la etiqueta para   *
*            la caja correspondiente          *
***********************************************
Parameters Vtel

if type("Vtel")<>"C"
 	Vtel = "E"
endif
do buscaiba
set refresh to 1
Store 0 to membarque,mbulto,indic,subin,mdesc,mdes1,mdes2
Store space(10) to mcliente,morsurped,mtienda
store 0 to morden
mcolor = ' '
cueco = ''
mleyenda=''
mtipo=''
MPZAS = 1
* ON KEY LABEL F1 DO PIEZAS
ON KEY 
close databases
@  4,1  clear to 20,78
@ 22,15 say space(63)
@  4,49 say 'Empaque para '+iif(Vtel="E","Exportacion",iif(Vtel="C","Centro de Negocios","Ventas x Telefono"))
@ 6,005 say 'Embarque  :'
@ 6,025 say 'Leyenda "Edr." (S/N) ? '
@ 6,050 say 'Tipo :'
@ 6,060 say 'Orden :'
@ 7,005 say 'Cliente   :'
@ 7,062 say 'Caja :'
@ 8,005 say 'Tienda    :'
@ 9,1 say replicate('Ä',78)
@ 10,2  say 'Clave' 
@ 10,31 say 'Descripcion'
@ 10,65 say 'Color'
@ 10,75 say 'Pza'
@ 11,1 say replicate('Ä',78)
ab = 'No.'
@ 21,5  get ab
ab = ' TOTAL '
@ 21,45 get ab
clear gets
do while .t.
 	use \clientes\quezada\exporta order embarque
 	store 0 to membarque
 	@ 22,15 say "[ESC] para salir [ENTER] nuevo embarque" color RB+/B
 	@ 6,017 get membarque pict "999999" 
 	READ
 	@ 22,15 say space(62)
 	set order to embarque descending
 	if membarque <> 0
  		seek STR(membarque,6,0)
  		if found()
   			store cvecliente to mcliente
   			store cvetienda  to mtienda
   			store leyenda    to mleyenda
   			store tipo       to mtipo
   			store orden      to morden 
  		else
   			mleyenda=''
   			mtipo=''
   			wait window "Embarque no existe ®Enter para continuar¯"
   			loop
   			use
  		endif
 	endif
 	use
 	if lastkey() = 27
  		exit
 	endif
 	if vtel="E"
  		if len(alltrim(mleyenda))=0
   			do while .t.
    			mleyenda = "N"
    			@ 6,048 get mleyenda pict "!"
    			read
    			if mleyenda $"SsNn"
     				exit
    			endif
   			enddo
  		else
   			@ 6,048 get mleyenda pict "!"
   			clear gets
  		endif
  		if len(alltrim(mtipo))=0
   			do while .t.
    			mtipo = " "
    			@ 22,15 say "[A]Autoservicios [T]Tiendas [E]Exportaciones [M]Mayoreo " ;
    						color RB+/B
    			@ 6,057 get mtipo pict "!"
    			read
    			@ 22,15 say space(62)
    			if mtipo $ "EeAaTtMm"
     				exit
    			endif
   			enddo
  		else
   			@ 6,057 get mtipo pict "!"
   			clear gets
  		endif
 	else
  		mleyenda="S"
  		mtipo=iif(vtel="V","V","C")
  		@ 6,048 get mleyenda pict "!"
  		@ 6,057 get mtipo pict "!"
  		clear gets
 	endif
 	@ 6,068 get morden pict "9999999999"
 	read
 	sele 1
 	use \access\inventa\ordensur order orden
 	use \clientes\quezada\pedidos order numpedido in select(1)
 	seek str(morden,10,0)
 	error = 0
 	if morden=0 or lastkey()=27
  		msg="Numero de Orden no debe ser Cero"
  		error = 1
 	else
  		if !found()
    		msg = "numero de orden no existe"
    		error = 1
  		else
   			if len(alltrim(mcliente))<>0
    			if cvecliente <> mcliente and cvetienda <> mtienda
     				msg = "Orden no pertenece al cliente y tienda del embarque"
     				store space(10) to mcliente
     				error = 1
    			endif
   			endif
   			store cvecliente to mcliente
   			store cvetienda  to mtienda
   			store numpedido  to mpedido
   			=seek(mpedido,"pedidos")
   			do case
    			case mtipo $ "CV"
     				if !vxtel 
      					msg="Orden de surtido no es de ventas por telefono"
      					store space(10) to mcliente
      					error=1
     				else
      					do case
        					case mtipo = "V" and tiposurt <>"V"
           						msg="Orden de surtido no es de ventas " + ;
           							"por telefono"
           						store space(10) to mcliente
           						error=1
        					case mtipo = "C" and tiposurt <>"C"
           						msg="Orden de surtido no es de Centro " + ;
           							"de Negocios"
           						store space(10) to mcliente
           						error=1
      					endcase
     				endif
    			case mtipo$"ET"
     				if !pedidos.tipoempres$"TE"
      					msg="Tipo de pedido no es de Exportaciones " + ;
      						"o Tiendas"
      					store space(10) to mcliente
      					error=1
     				endif
    			case mtipo="A"
     				if !pedidos.tipoempres$"A"
      					msg="Tipo de pedido no es de Autoservicios "
      					store space(10) to mcliente
      					error=1
     				endif
    			case mtipo="M"
     				if !pedidos.tipoempres$"N"
      					msg="Tipo de pedido no es de Mayoreo"
      					store space(10) to mcliente
      					error=1
     				else
      					if pedidos.tipodocto = "N"
       						msg="Solo aquellos pedidos de mayoreo " + ;
       							"que se Resurten"
       						store space(10) to mcliente
       						error=1
      					endif
     				endif
   			endcase
   			store pedidos.orden to morsurped
  		endif
 	endif
 	close databases
 	if error = 1
  		mleyenda=''
  		mtipo=''
  		wait window msg + "®Enter para continuar¯"
  		loop
 	endif
 	do case
  		case !VTEL$"VC"
    		use \clientes\quezada\clientes order cvecliente
  		case VTEL$"V"
    		use \access\vtel\clientt1 index \access\vtel\clientt1
  		case VTEL$"C"
    		use \access\cneg\clientt1 index \access\cneg\clientt1 
 	endcase
 	seek mcliente
 	store iif(vtel$"CV",nombre,razonsocia) to mrazon
 	use
 	@ 7,017 say mcliente +" "+left(mrazon,30)
 	
 	use \clientes\quezada\tiendas order cvetienda
 	seek mtienda
 	store direccion to mnomtienda
 	use
 	@ 08,017 say mtienda + " "+left(mnomtienda,40)
 	
 	***************************************************
	*   ARCHIVO DE CONTROL PARA BLOQUEOS DE EQUIPO    *
	***************************************************
	
 	mreconstruye = .f.
 	use c:\access\inventa\control
 	if reccount()<>0
  		if Tarea <> IIF(VTEL$"CV","V.3.1 ","X.1 ")
   			wait window "Dejo un proceso pendiente que no es Empaque"
			use
   			return
  		endif
  		nomfil = alltrim(modulo)
  		use (nomfil) excl
  		if reccount() <> 0
   			store orden to numcot
   			if numcot <> morden
    			wait window "Orden de Surtido leida anteriormente ;
    						no es la misma"
    			close databases
    			return
   			endif
   			mreconstruye = .t.
  		else
   			use c:\access\inventa\control excl
   			dele all
   			pack
   			use
  		endif
 	endif
 	use
	
 	sele 1
 	traspa = 'C:\access\inventa\' + sys(3)
 	if mreconstruye
    	traspa = nomfil
 	else
   		do provic with traspa
   		use &traspa exclu
   		index on clapr to &traspa
   		use
   		use c:\access\inventa\control
   		append blank
   		replace modulo with traspa, ;
   				Tarea with IIF(VTEL$"CV","V.3.1","X.1")
   		use
 	endif
 	
 	use \clientes\quezada\estilos  order cveestilo           in 2
 	use \clientes\quezada\colores  order cvecolor            in 3
 	use \clientes\quezada\exporta  order embarque descending in 4
 	
	
 	msalir = .f.
 	re = 12
 	Store space(15) to claux,clant
 	do while !msalir
  		select 1
  		use
  		if re = 20
   			re = 11
  		endif
 		
  		ON KEY LABEL F1 DO saldos
 		
  		Store space(15) to claux
  		@  re,1  clear to 20,78
  		@ re,001 get claux picture '@!'
  		read
		
  		ON KEY
  		
  		sele 1
  		use &traspa index &traspa exclusive
  		if len(alltrim(claux))=0
    		espera = ' '
    		do while .not. espera $ 'TC'
     			do mensaj with 13, 'S'
    		enddo
    		if espera = 'T'
     			if reccount()=0
      				msalir = .t.
     			else
      				espera = ' '
      				do while .not. espera $ 'SN'
       					do mensaj with 8, 'S'
      				enddo
      				if espera <> 'S'
       					sele 1
       					zap
       					use
      				else
       					do cajas
      				endif
      				@  re,1  clear to 20,78
      				RE = RE + 1
      				indic = 0
     			endif
    		endif
    		select 1
    		use
    		loop
  		endif
		
		claux = ltrim(claux)
  		use \access\inventa\codprov order proveedor in select(1)
  		select codprov
  		encontro=seek(claux,"CODPROV")
  		if encontro
     		store alltrim(codprov.cveestilo) to claux,cla01
     		store codprov.color              to mcolor
     		store val(cla01) to cla01
  		else   
     		claux = alltrim(claux)
     		cla01 = val(substr(claux, 1, iif(len(claux) > 5, 6, 5)))
     		claux = substr(claux, iif(len(claux) > 5, 7, 1), 5)
  		endif
  		use
  		
  		sele estilos
  		seek claux
  		mlinea = " "
  		if found()
   			? chr(7),chr(7)
   			@ re,1  say space(15)
   			@ re,2  say claux
   			@ 22,15 say space(63)
   			mlinea = estilos.linea
   			apunt = 1
   			colum = 19
   			cueco = ''
   			mpzas = iif(cla01 <= 10, cla01, 1)
   			do while apunt <= tallas
    			camau = 'talla' + str(apunt, 1, 0)
    			camau = &camau
    			select colores
    			seek camau
    			@ 22,colum say '(' + str(apunt, 1, 0) + ')' + abrev
    			select estilos
    			cueco = cueco + str(apunt, 1, 0)
    			apunt = apunt + 1
    			colum = colum + 7
   			enddo
   			@ re,20  say descripcio
   			read
   			if encontro
     			@ re,67 say mcolor picture '9' 
   			else
     			if apunt > 2
       				@ re,67 get mcolor picture '9' valid(mcolor $ cueco) 
       				if claux <> clant
        				?? chr(7),chr(7)   
        				read
       				endif
       				clear gets
     			else
      				mcolor = '1'
      				@ re,67 say mcolor picture '9' 
     			endif
   			endif
   			@  re,73 get mpzas pict '9,999'
   			if mlinea = "90        " or mlinea = "014       "
    			read
   			endif
   			clear gets
   			use \access\inventa\ordensur   order ordenesti  in 5
   			use \clientes\quezada\detapedi order estnumpedi in 6
   			sele ordensur
   			mkey = str(morden,10,0)+claux+replicate(" ",10-len(claux))
   			seek mkey
   			if !found()
    			?? chr(7),chr(7),chr(7),chr(7)
    			do mensaj with 7, 'S'
    			wait window "Articulo no existe en Orden de Surtido"
    			use
    			loop
   			endif
   			mkey = claux+replicate(" ",10-len(claux))+mpedido
   			=seek(mkey,"detapedi")
   			if !found()
                espera = ' '
                do while not espera $ 'Cc'
                	msg = '>>> ARTICULO  ' + claux + ' No se Encontro '+;
                	      'en el Pedido ' + mpedido + ' <<<' + chr(13) + ;
                	      'llame a sistemas para verificar el dato y' + ;
                	      'corregir el problema en la Orden'
                	DO ALERTA with msg
                enddo
                loop
   			endif
   			store detapedi.candevu1 to mdesc
   			store detapedi.candevu2 to mdes1
   			store detapedi.candevu3 to mdes2
            mpe = 'cant' + mcolor
            msa = 'surt' + mcolor
            
            sum &mpe,&msa to tot,msurt ;
            	while orden = str(morden,10) and cveestilo = claux
            
   			mexi = tot - msurt
   			sele 5
   			use
   			sele 6
   			use
   			select 1
   			seek claux
   			canpr = 'canp' + mcolor
   			mexi = mexi - &canpr
   			
   			if mexi < mpzas
    			?? chr(7),chr(7),chr(7),chr(7)
    			do mensaj with 7, 'S'
    			wait window "Existencia para Pedido : "+str(mexi,12,2)
    			loop
   			endif
   			
   			select 1
   			seek claux
   			
   			if .not. found()
    			append blank
    			replace clapr with claux, ;
    					despr with mdesc, ;
    					orden with morden
   			endif
   			canpr = 'canp' + mcolor
 			
   			replace &canpr with (&canpr + mpzas)
   			use
   			if claux <> clant
    			if clant <> space(15)
     				subin = 0
    			endif
   			endif
   			indic = indic + mpzas
   			subin = subin + mpzas
   			mtmp = reccount(1)
   			@ 21,10    get indic picture '###'
   			@ 21,18    get subin picture '###'
   			@ 21,55    get mtmp picture '###'
   			clear gets
   			re = re + 1
   			clant = claux
   			store space(15) to claux
  		else
   			clant = claux
   			store space(15) to claux
   			? chr(7),chr(7),chr(7),chr(7),chr(7),chr(7)
   			do mensaj with 7, 'S'
   			select 1
   			use
   			loop
  		endif
 	enddo
 	exit
enddo
@  4,1  clear to 20,78
use c:\access\inventa\control excl
zap
use
close databases
dele file &traspa".dbf"
dele file &traspa".cdx"
ON KEY 
return

Procedure Cajas
sele 1
use &traspa index &traspa exclusive
use \access\inventa\ordensur   order ordenesti  in 5
use \clientes\quezada\detapedi order estnumpedi in 6

mtmp1 = "\access\temporal\"+ sys(3)+".fmt"
mtmp2 = "\access\temporal\"+ sys(3)+".bch"
mtmp3 = "\access\temporal\"+ sys(3)+".fmt"
mtmp4 = "\access\temporal\"+ sys(3)+".bch"
mtmp5 = "\access\temporal\"+ sys(3)+".fmt"
mtmp6 = "\access\temporal\"+ sys(3)+".bch"
mtmp7 = "\access\temporal\"+ sys(3)+".fmt"
mtmp8 = "\access\temporal\"+ sys(3)+".bch"
mtmp9 = "\access\temporal\"+ sys(3)+".fmt"
mtmp10= "\access\temporal\"+ sys(3)+".bch"
mtmp11= "\access\temporal\"+ sys(3)+".fmt"
mtmp12= "\access\temporal\"+ sys(3)+".bch"
mheader1 = fcreate(mtmp1)
mheader2 = fcreate(mtmp2)
mheader3 = fcreate(mtmp3)
mheader4 = fcreate(mtmp4)
mheader5 = fcreate(mtmp5)
mheader6 = fcreate(mtmp6)
mheader7 = fcreate(mtmp7)
mheader8 = fcreate(mtmp8)
mheader9 = fcreate(mtmp9)
mheader10= fcreate(mtmp10)
mheader11= fcreate(mtmp11)
mheader12= fcreate(mtmp12)
store "" to ma1,ma2,ma3,ma4,ma5,ma6,ma7,ma8,ma9,ma10, ;
 			ma11,ma12,ma13,ma14,ma15,ma16,ma17,ma18,ma19,ma20, ;
 			ma21,ma22,ma23,ma24,ma25,ma26,ma27,ma28,ma29,ma30
store "" to ma31,ma32,ma33,ma34,ma35,ma36,ma37,ma38,ma39,ma40, ;
 			ma41,ma42,ma43,ma44,ma45,ma46,ma47,ma48,ma49,ma50, ;
 			ma51,ma52,ma53,ma54,ma55,ma56,ma57,ma58,ma59,ma60
go top
mcnt = 1
primero = .t.
scan
 	store clapr to mesti
 	store canp1 to mpza1
 	store canp2 to mpza2
 	store canp3 to mpza3
 	store canp4 to mpza4
 	store canp5 to mpza5
 	store canp6 to mpza6
 	store despr to mdesc
 	sele estilos
 	seek mesti
 	do while !rlock()
 	enddo
 	store descripcio to mdes
 	mcol1 = iif(mpza1>0,iif(seek(estilos.talla1,"colores"),colores.abrev,""),"")
 	mcol2 = iif(mpza2>0,iif(seek(estilos.talla2,"colores"),colores.abrev,""),"")
 	mcol3 = iif(mpza3>0,iif(seek(estilos.talla3,"colores"),colores.abrev,""),"")
 	mcol4 = iif(mpza4>0,iif(seek(estilos.talla4,"colores"),colores.abrev,""),"")
 	mcol5 = iif(mpza5>0,iif(seek(estilos.talla5,"colores"),colores.abrev,""),"")
 	mcol6 = iif(mpza6>0,iif(seek(estilos.talla6,"colores"),colores.abrev,""),"")
 	mva = "ma"+alltrim(str(mcnt,2,0))
	&mva = &mva + left(alltrim(mdes),35)
 	for i = 1 to 6
  		mcan = "mpza"+str(i,1,0)
  		mcol = "mcol"+str(i,1,0)
  		if &mcan > 0
   			mstr = " "+alltrim(str(&mcan))+' '+&mcol
   			if len(mva)+len(mstr) > 26
    			mcnt = mcnt + 1
    			mva = "ma"+alltrim(str(mcnt,2,0))
   			endif
   			&mva = &mva + mstr
  		endif
 	endfor
 	mcnt = mcnt + 1
 	sele exporta
 	if primero 
  		if membarque = 0
   			do while !flock()
   			enddo
   			set order to embarque descending
   			go top
   			store embarque + 1 to membarque
   			@ 6,017 get membarque pict "999999" 
  		endif
  		seek STR(membarque,6,0)
  		do while !flock()
  		enddo
  		store bulto + 1  to mbulto
  		@ 7,069 get mbulto picture "99999"
  		clear gets
  		primero = .f.
 	endif
 	mpl=0
 	if !mtipo$"ECV"
  		do case
   			case mtipo = "A"
    			mpl = round(estilos.preciolist*1.05,0)
   			otherwise
    			mpl = estilos.preciolist
  		endcase
 	endif
 	
 	sele ordensur
 	mkey = str(morden,10,0)+mesti+replicate(" ",10-len(mesti))
 	seek mkey
 	msurt = 0
	scan while orden = str(morden,10,0) and cveestilo = mesti
   		store numpedido to mpedido
 		mc1 = iif((cant1 -surt1)>mpza1,mpza1,cant1-surt1)
		mc2 = iif((cant2 -surt2)>mpza2,mpza2,cant2-surt2)
    	mc3 = iif((cant3 -surt3)>mpza3,mpza3,cant3-surt3)
    	mc4 = iif((cant4 -surt4)>mpza4,mpza4,cant4-surt4)
    	mc5 = iif((cant5 -surt5)>mpza5,mpza5,cant5-surt5)
    	mc6 = iif((cant6 -surt6)>mpza6,mpza6,cant6-surt6)
		
		msurt = msurt + mc1+mc2+mc3+mc4+mc5+mc6
		mpza1 = mpza1 - mc1
		mpza2 = mpza2 - mc2
		mpza3 = mpza3 - mc3
		mpza4 = mpza4 - mc4
		mpza5 = mpza5 - mc5
		mpza6 = mpza6 - mc6
		do while !rlock()
		enddo
		
 		sele detapedi
 		mkey = mesti+replicate(" ",10-len(mesti))+mpedido
 		seek mkey
		do while !rlock()
		enddo
		
 		store mc1 to mcant1
 		store mc2 to mcant2
 		store mc3 to mcant3
 		store mc4 to mcant4
 		store mc5 to mcant5
 		store mc6 to mcant6
 		mtotal = mcant1 + mcant2 + mcant3 + mcant4 + mcant5 + mcant6
		do while mtotal > 0
			sele estilos
			exi2004 = (exifin04_1 + exifin04_2 + exifin04_3 + exifin04_4 + ;
                    exifin04_5 + exifin04_6) + (entra04_1 + entra04_2 + ;
					entra04_3 + entra04_4 + entra04_5 + entra04_6) - ;
					(salida04_1 + salida04_2 + salida04_3 + salida04_4 + ;
					salida04_5 + salida04_6)
			mexi1 = iif(exifin04_1+entra04_1-salida04_1<0, 0, ;
					exifin04_1+entra04_1-salida04_1)
			mexi2 = iif(exifin04_2+entra04_2-salida04_2<0, 0, ;
					exifin04_2+entra04_2-salida04_2)
			mexi3 = iif(exifin04_3+entra04_3-salida04_3<0, 0, ;
					exifin04_3+entra04_3-salida04_3)
			mexi4 = iif(exifin04_4+entra04_4-salida04_4<0, 0, ;
					exifin04_4+entra04_4-salida04_4)
			mexi5 = iif(exifin04_5+entra04_5-salida04_5<0, 0, ;
					exifin04_5+entra04_5-salida04_5)
			mexi6 = iif(exifin04_6+entra04_6-salida04_6<0, 0, ;
					exifin04_6+entra04_6-salida04_6)
			
			if exi2004 > 0
				manocosto = "2004"
			else
				manocosto = "2005"
			endif
			store 0 to mct1,mct2,mct3,mct4,mct5,mct6
			if manocosto = "2004"
				mct1=iif(mexi1>mcant1, mcant1, mexi1)
			else
				mct1 = mcant1
			endif
			if manocosto = "2004" 
				mct2=iif(mexi2>mcant2, mcant2, mexi2)
			else
				mct2 = mcant2
			endif
			if manocosto = "2004" 
				mct3=iif(mexi3>mcant3, mcant3, mexi3)
			else
				mct3 = mcant3
			endif
			if manocosto = "2004" 
				mct4=iif(mexi4>mcant4, mcant4, mexi4)
			else
				mct4 = mcant4
			endif
			if manocosto = "2004" 
				mct5=iif(mexi5>mcant5, mcant5, mexi5)
			else
				mct5 = mcant5
			endif
			if manocosto = "2004" 
				mct6=iif(mexi6>mcant6, mcant6, mexi6)
			else
				mct6 = mcant6
			endif	
			if manocosto = "2005"
        		mcosafecta = "estilos.costo"
				mmvto1 = "estilos.surt1"
                mmvto2 = "estilos.surt2"
                mmvto3 = "estilos.surt3"
                mmvto4 = "estilos.surt4"
                mmvto5 = "estilos.surt5"
                mmvto6 = "estilos.surt6"
			else
        		mcosafecta = "estilos.costo2004"
				mmvto1 = "estilos.salida04_1"
				mmvto2 = "estilos.salida04_2"
				mmvto3 = "estilos.salida04_3"
				mmvto4 = "estilos.salida04_4"
				mmvto5 = "estilos.salida04_5"
				mmvto6 = "estilos.salida04_6"
			endif
			
		
			select exporta
			********************
			* exporta          *
			********************
			
 			append blank
 			if detapedi.orden = str(morden,10,0)
 				replace exporta.embarque   with membarque, ;
 						exporta.bulto      with mbulto, ;
 						exporta.cvecliente with mcliente, ;
 						exporta.cvetienda  with mtienda, ;
 						exporta.tipo       with mtipo, ;
	 					exporta.leyenda    with mleyenda, ;
 						exporta.fechaalta  with date(), ;
 						exporta.horamvto   with time(), ;
 						exporta.orden      with morden, ;
 						exporta.numpedido  with mpedido, ;
	 					exporta.cveestilo  with mesti, ;
 						exporta.cant1      with mct1, ;
 						exporta.cant2      with mct2, ;
 						exporta.cant3      with mct3, ;
 						exporta.cant4      with mct4, ;
	 					exporta.cant5      with mct5, ;
 						exporta.cant6      with mct6, ;
  						exporta.preciolist with mpl, ;
 						exporta.costo      with &mcosafecta, ;
 						exporta.anocosto   with manocosto, ;
	 					exporta.status     with '1'
 				replace detapedi.cansurt1  with detapedi.cansurt1 + mct1, ;
 						detapedi.cansurt2  with detapedi.cansurt2 + mct2, ;
 						detapedi.cansurt3  with detapedi.cansurt3 + mct3, ;
 						detapedi.cansurt4  with detapedi.cansurt4 + mct4, ;
	 					detapedi.cansurt5  with detapedi.cansurt5 + mct5, ;
 						detapedi.cansurt6  with detapedi.cansurt6 + mct6, ;
 						detapedi.pedsurt1  with detapedi.pedsurt1 + mct1, ;
 						detapedi.pedsurt2  with detapedi.pedsurt2 + mct2, ;
 						detapedi.pedsurt3  with detapedi.pedsurt3 + mct3, ;
 						detapedi.pedsurt4  with detapedi.pedsurt4 + mct4, ;
	 					detapedi.pedsurt5  with detapedi.pedsurt5 + mct5, ;
 						detapedi.pedsurt6  with detapedi.pedsurt6 + mct6, ;
 						ordensur.fcierre   with date(), ;
 						ordensur.cerrado   with .f., ;
 						ordensur.surt1     with ordensur.surt1 + mct1, ;
 						ordensur.surt2     with ordensur.surt2 + mct2, ;
 						ordensur.surt3     with ordensur.surt3 + mct3, ;
	 					ordensur.surt4     with ordensur.surt4 + mct4, ;
 						ordensur.surt5     with ordensur.surt5 + mct5, ;
 						ordensur.surt6     with ordensur.surt6 + mct6
 				replace &mmvto1      with &mmvto1 + mct1, ;
 						&mmvto2      with &mmvto2 + mct2, ;
	 					&mmvto3      with &mmvto3 + mct3, ;
 						&mmvto4      with &mmvto4 + mct4, ;
 						&mmvto5      with &mmvto5 + mct5, ;
 						&mmvto6      with &mmvto6 + mct6
			else
 				replace exporta.embarque   with membarque, ;
 						exporta.bulto      with mbulto, ;
 						exporta.cvecliente with mcliente, ;
 						exporta.cvetienda  with mtienda, ;
 						exporta.tipo       with mtipo, ;
	 					exporta.leyenda    with mleyenda, ;
 						exporta.fechaalta  with date(), ;
 						exporta.horamvto   with time(), ;
 						exporta.orden      with morden, ;
 						exporta.numpedido  with mpedido, ;
	 					exporta.cveestilo  with mesti, ;
 						exporta.cant1      with mct1, ;
 						exporta.cant2      with mct2, ;
 						exporta.cant3      with mct3, ;
 						exporta.cant4      with mct4, ;
	 					exporta.cant5      with mct5, ;
 						exporta.cant6      with mct6, ;
  						exporta.preciolist with mpl, ;
 						exporta.costo      with &mcosafecta, ;
 						exporta.anocosto   with manocosto, ;
	 					exporta.status     with '1'
 				replace detapedi.cansurt1  with detapedi.cansurt1 + mct1, ;
 						detapedi.cansurt2  with detapedi.cansurt2 + mct2, ;
 						detapedi.cansurt3  with detapedi.cansurt3 + mct3, ;
 						detapedi.cansurt4  with detapedi.cansurt4 + mct4, ;
	 					detapedi.cansurt5  with detapedi.cansurt5 + mct5, ;
 						detapedi.cansurt6  with detapedi.cansurt6 + mct6, ;
 						ordensur.fcierre   with date(), ;
 						ordensur.cerrado   with .f., ;
 						ordensur.surt1     with ordensur.surt1 + mct1, ;
 						ordensur.surt2     with ordensur.surt2 + mct2, ;
 						ordensur.surt3     with ordensur.surt3 + mct3, ;
	 					ordensur.surt4     with ordensur.surt4 + mct4, ;
 						ordensur.surt5     with ordensur.surt5 + mct5, ;
 						ordensur.surt6     with ordensur.surt6 + mct6, ;
 						&mmvto1      with &mmvto1 + mct1, ;
 						&mmvto2      with &mmvto2 + mct2, ;
	 					&mmvto3      with &mmvto3 + mct3, ;
 						&mmvto4      with &mmvto4 + mct4, ;
 						&mmvto5      with &mmvto5 + mct5, ;
 						&mmvto6      with &mmvto6 + mct6
			endif
 			flush
			mtotal = mtotal - (mct1+mct2+mct3+mct4+mct5+mct6)
			mcant1 = mcant1 - mct1
			mcant2 = mcant2 - mct2
			mcant3 = mcant3 - mct3
			mcant4 = mcant4 - mct4
			mcant5 = mcant5 - mct5
			mcant6 = mcant6 - mct6
		enddo	
 		
 		unlock all
 		
 		use \clientes\quezada\exporta  order embarque descending 
		
 		sele detapedi
 		
 		set order to numpedido
 		seek mpedido
 		replace fcierre with date() while numpedido = mpedido
 		
 		use \clientes\quezada\detapedi order estnumpedi
 		
 		sele estilos
 		unlock
 		use \clientes\quezada\estilos  order cveestilo
 		
	 	select select(1)
 		use \clientes\quezada\pedidos order numpedido
 		seek mpedido
		if mtipo $ "CV"
 			replace fcierre with date(), ;
 					plazo with 1 ;
 			 		while numpedido = mpedido
 		else
 			replace fcierre with date() ;
 					while numpedido = mpedido
		endif
 		use
		
 		sele ordensur
		unlock
		if (mpza1+mpza2+mpza3+mpza4+mpza5+mpza6) <= 0
	   		exit
		endif
	endscan
	sele ordensur
	use \access\inventa\ordensur order ordenesti
 	sele 1
endscan
zap
use
sele detapedi
use
sele ordensur
use
select 1
MCNT = MCNT - 1
 	ML1=ALLTRIM(STR(MEMBARQUE,6,0))
 	ML1= REPLICATE('0',6-LEN(ML1))+ML1
 	ML2=ALLTRIM(STR(MBULTO,5,0))
 	ML2= REPLICATE('0',5-LEN(ML2))+ML2
 	=fputs(mheader1,'US')
 	=fputs(mheader1,'N')
 	=fputs(mheader1,'Q810,24')
 	=fputs(mheader1,'q816')
 	=fputs(mheader1,'S2')
 	=fputs(mheader1,'D8')
 	=fputs(mheader1,'ZT')
 	=fputs(mheader1,'A812,0,1,1,6,1,N,"&ma1"')
 	=fputs(mheader1,'A451,6,1,1,6,1,N,"&ma6"')
 	=fputs(mheader1,'A544,6,1,1,6,1,N,"&ma5"')
 	=fputs(mheader1,'A363,8,1,1,6,1,N,"&ma3"')
 	=fputs(mheader1,'A637,8,1,1,6,1,N,"&ma2"')
 	=fputs(mheader1,'A727,8,1,1,6,1,N,"&ma4"')
 	=fputs(mheader1,'A158,12,1,5,1,1,N,"CAJA     :"')
 	=fputs(mheader1,'A270,12,1,5,1,1,N,"EMBARQUE :"')
 	=fputs(mheader1,'A80,0,1,4,3,1,N,"&mrazon"')
 	=fputs(mheader1,'A199,398,1,5,2,1,N,"&ML2"')
 	=fputs(mheader1,'A290,398,1,5,2,1,N,"&ML1"')
 	=fputs(mheader1,'B320,730,0,UA0,4,4,122,B,"&ML1&ML2"')
 	=fputs(mheader1,'P1')
 	IF MCNT > 6
  		=fputs(mheader2,'US')
  		=fputs(mheader2,'N')
  		=fputs(mheader2,'Q810,24')
  		=fputs(mheader2,'q816')
  		=fputs(mheader2,'S2')
  		=fputs(mheader2,'D8')
  		=fputs(mheader2,'ZT')
  		=fputs(mheader2,'A812,0,1,1,6,1,N,"&ma7"')
  		=fputs(mheader2,'A451,6,1,1,6,1,N,"&ma12"')
  		=fputs(mheader2,'A544,6,1,1,6,1,N,"&ma11"')
  		=fputs(mheader2,'A363,8,1,1,6,1,N,"&ma9"')
  		=fputs(mheader2,'A637,8,1,1,6,1,N,"&ma8"')
  		=fputs(mheader2,'A727,8,1,1,6,1,N,"&ma10"')
  		=fputs(mheader2,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader2,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader2,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader2,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader2,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader2,'P1')
 	ENDIF
 	IF MCNT > 12
  		=fputs(mheader3,'US')
  		=fputs(mheader3,'N')
  		=fputs(mheader3,'Q810,24')
  		=fputs(mheader3,'q816')
  		=fputs(mheader3,'S2')
  		=fputs(mheader3,'D8')
  		=fputs(mheader3,'ZT')
  		=fputs(mheader3,'A812,0,1,1,6,1,N,"&ma13"')
  		=fputs(mheader3,'A451,6,1,1,6,1,N,"&ma18"')
  		=fputs(mheader3,'A544,6,1,1,6,1,N,"&ma17"')
  		=fputs(mheader3,'A363,8,1,1,6,1,N,"&ma15"')
  		=fputs(mheader3,'A637,8,1,1,6,1,N,"&ma14"')
  		=fputs(mheader3,'A727,8,1,1,6,1,N,"&ma16"')
  		=fputs(mheader3,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader3,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader3,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader3,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader3,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader3,'P1')
 	ENDIF
 	IF MCNT > 18
  		=fputs(mheader4,'US')
  		=fputs(mheader4,'N')
  		=fputs(mheader4,'Q810,24')
  		=fputs(mheader4,'q816')
  		=fputs(mheader4,'S2')
  		=fputs(mheader4,'D8')
  		=fputs(mheader4,'ZT')
  		=fputs(mheader4,'A812,0,1,1,6,1,N,"&ma19"')
  		=fputs(mheader4,'A451,6,1,1,6,1,N,"&ma24"')
  		=fputs(mheader4,'A544,6,1,1,6,1,N,"&ma23"')
  		=fputs(mheader4,'A363,8,1,1,6,1,N,"&ma21"')
  		=fputs(mheader4,'A637,8,1,1,6,1,N,"&ma20"')
  		=fputs(mheader4,'A727,8,1,1,6,1,N,"&ma22"')
  		=fputs(mheader4,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader4,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader4,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader4,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader4,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader4,'P1')
 	ENDIF
 	IF MCNT > 24
  		=fputs(mheader5,'US')
  		=fputs(mheader5,'N')
  		=fputs(mheader5,'Q810,24')
  		=fputs(mheader5,'q816')
  		=fputs(mheader5,'S2')
  		=fputs(mheader5,'D8')
  		=fputs(mheader5,'ZT')
  		=fputs(mheader5,'A812,0,1,1,6,1,N,"&ma25"')    &&1
  		=fputs(mheader5,'A451,6,1,1,6,1,N,"&ma30"')    &&6
  		=fputs(mheader5,'A544,6,1,1,6,1,N,"&ma29"')    &&5
  		=fputs(mheader5,'A363,8,1,1,6,1,N,"&ma27"')    &&3
  		=fputs(mheader5,'A637,8,1,1,6,1,N,"&ma26"')    &&2
  		=fputs(mheader5,'A727,8,1,1,6,1,N,"&ma28"')    &&4
  		=fputs(mheader5,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader5,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader5,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader5,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader5,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader5,'P1')
 	ENDIF
 	IF MCNT > 30
  		=fputs(mheader6,'US')
  		=fputs(mheader6,'N')
  		=fputs(mheader6,'Q810,24')
  		=fputs(mheader6,'q816')
  		=fputs(mheader6,'S2')
  		=fputs(mheader6,'D8')
  		=fputs(mheader6,'ZT')
  		=fputs(mheader6,'A812,0,1,1,6,1,N,"&ma31"')       &&1
  		=fputs(mheader6,'A451,6,1,1,6,1,N,"&ma36"')       &&6
  		=fputs(mheader6,'A544,6,1,1,6,1,N,"&ma35"')       &&5
  		=fputs(mheader6,'A363,8,1,1,6,1,N,"&ma33"')       &&3
  		=fputs(mheader6,'A637,8,1,1,6,1,N,"&ma32"')       &&2
  		=fputs(mheader6,'A727,8,1,1,6,1,N,"&ma34"')       &&4
  		=fputs(mheader6,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader6,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader6,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader6,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader6,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader6,'P1')
 	ENDIF
 	IF MCNT > 36
  		=fputs(mheader7,'US')
  		=fputs(mheader7,'N')
  		=fputs(mheader7,'Q810,24')
  		=fputs(mheader7,'q816')
  		=fputs(mheader7,'S2')
  		=fputs(mheader7,'D8')
  		=fputs(mheader7,'ZT')
  		=fputs(mheader7,'A812,0,1,1,6,1,N,"&ma37"')       &&1
  		=fputs(mheader7,'A451,6,1,1,6,1,N,"&ma42"')       &&6
  		=fputs(mheader7,'A544,6,1,1,6,1,N,"&ma41"')       &&5
  		=fputs(mheader7,'A363,8,1,1,6,1,N,"&ma39"')       &&3
  		=fputs(mheader7,'A637,8,1,1,6,1,N,"&ma38"')       &&2
  		=fputs(mheader7,'A727,8,1,1,6,1,N,"&ma40"')       &&4
  		=fputs(mheader7,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader7,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader7,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader7,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader7,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader7,'P1')
 	ENDIF
 	IF MCNT > 42
  		=fputs(mheader8,'US')
  		=fputs(mheader8,'N')
  		=fputs(mheader8,'Q810,24')
  		=fputs(mheader8,'q816')
  		=fputs(mheader8,'S2')
  		=fputs(mheader8,'D8')
  		=fputs(mheader8,'ZT')
  		=fputs(mheader8,'A812,0,1,1,6,1,N,"&ma43"')       &&1
  		=fputs(mheader8,'A451,6,1,1,6,1,N,"&ma48"')       &&6
  		=fputs(mheader8,'A544,6,1,1,6,1,N,"&ma47"')       &&5
  		=fputs(mheader8,'A363,8,1,1,6,1,N,"&ma45"')       &&3
  		=fputs(mheader8,'A637,8,1,1,6,1,N,"&ma44"')       &&2
  		=fputs(mheader8,'A727,8,1,1,6,1,N,"&ma46"')       &&4
  		=fputs(mheader8,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader8,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader8,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader8,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader8,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader8,'P1')
 	ENDIF
 	IF MCNT > 48
  		=fputs(mheader9,'US')
  		=fputs(mheader9,'N')
  		=fputs(mheader9,'Q810,24')
  		=fputs(mheader9,'q816')
  		=fputs(mheader9,'S2')
  		=fputs(mheader9,'D8')
  		=fputs(mheader9,'ZT')
  		=fputs(mheader9,'A812,0,1,1,6,1,N,"&ma49"')       &&1
  		=fputs(mheader9,'A451,6,1,1,6,1,N,"&ma54"')       &&6
  		=fputs(mheader9,'A544,6,1,1,6,1,N,"&ma53"')       &&5
  		=fputs(mheader9,'A363,8,1,1,6,1,N,"&ma51"')       &&3
  		=fputs(mheader9,'A637,8,1,1,6,1,N,"&ma50"')       &&2
  		=fputs(mheader9,'A727,8,1,1,6,1,N,"&ma52"')       &&4
  		=fputs(mheader9,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader9,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader9,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader9,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader9,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader9,'P1')
 	ENDIF
 	IF MCNT > 54
  		=fputs(mheader10,'US')
  		=fputs(mheader10,'N')
  		=fputs(mheader10,'Q810,24')
  		=fputs(mheader10,'q816')
  		=fputs(mheader10,'S2')
  		=fputs(mheader10,'D8')
  		=fputs(mheader10,'ZT')
  		=fputs(mheader10,'A812,0,1,1,6,1,N,"&ma55"')       &&1
  		=fputs(mheader10,'A451,6,1,1,6,1,N,"&ma60"')       &&6
  		=fputs(mheader10,'A544,6,1,1,6,1,N,"&ma59"')       &&5
  		=fputs(mheader10,'A363,8,1,1,6,1,N,"&ma57"')       &&3
  		=fputs(mheader10,'A637,8,1,1,6,1,N,"&ma56"')       &&2
  		=fputs(mheader10,'A727,8,1,1,6,1,N,"&ma58"')       &&4
  		=fputs(mheader10,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader10,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader10,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader10,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader10,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader10,'P1')
 	ENDIF
=fclose(mheader1)
=fclose(mheader2)
=fclose(mheader3)
=fclose(mheader4)
=fclose(mheader5)
=fclose(mheader6)
=fclose(mheader7)
=fclose(mheader8)
=fclose(mheader9)
=fclose(mheader10)
=fclose(mheader11)
=fclose(mheader12)
IF mtipo$"TAEVCM"
 	COM="LPT1:"
ELSE
 	COM="COM1:"
ENDIF
IF !mtipo$"TAEVC"
 	! COPY &MTMP1 &COM > c:\access\inventa\TXT.txt
 	WAIT WINDOW "ENTER SIGUIENTE ARCHIVO"
 	! COPY &MTMP2 &COM > c:\access\inventa\TXT.txt
 	IF MCNT > 3
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP3 &COM > c:\access\inventa\TXT.txt
  		WAIT WINDOW "ENTER SIGUIENTE ARCHIVO"
  		! COPY &MTMP4 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 6
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP5 &COM > c:\access\inventa\TXT.txt
  		WAIT WINDOW "ENTER SIGUIENTE ARCHIVO"
  		! COPY &MTMP6 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 9
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP7 &COM > c:\access\inventa\TXT.txt
  		WAIT WINDOW "ENTER SIGUIENTE ARCHIVO"
  		! COPY &MTMP8 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 12
 		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP9 &COM > c:\access\inventa\TXT.txt
  		WAIT WINDOW "ENTER SIGUIENTE ARCHIVO"
  		! COPY &MTMP10 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 15
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP11 &COM > c:\access\inventa\TXT.txt
  		WAIT WINDOW "ENTER SIGUIENTE ARCHIVO"
  		! COPY &MTMP12 &COM > c:\access\inventa\TXT.txt
 	ENDIF
ELSE
 	! COPY &MTMP1 &COM > c:\access\inventa\TXT.txt
 	IF MCNT > 6
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP2 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 12
 		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP3 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 18
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP4 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 24
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP5 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 30
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP6 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 36
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP7 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 42
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP8 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 48
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP9 &COM > c:\access\inventa\TXT.txt
 	ENDIF
 	IF MCNT > 54
  		WAIT WINDOW "ENTER SIGUIENTE ETIQUETA"
  		FOR I = 1 TO 1000
  		ENDFOR
  		! COPY &MTMP10 &COM > c:\access\inventa\TXT.txt
 	ENDIF
ENDIF
dele file &mtmp1
dele file &mtmp2
dele file &mtmp3
dele file &mtmp4
dele file &mtmp5
dele file &mtmp6
dele file &mtmp7
dele file &mtmp8
dele file &mtmp9
dele file &mtmp10
dele file &mtmP11
dele file &mtmp12
return

procedure piezas
define window piezas from 10,20 to 20,60
activate window piezas
IF LEN(CUECO)=0
 	CUECO = '1'
ENDIF
do while .t.
 	@ 3,10 say 'Color  :' get mcolor pict '9'
 	IF LEN(CUECO)=1
  		CLEAR GETS
  		EXIT
 	ELSE
  		read
 	ENDIF
 	if mcolor $ cueco
  		exit
 	endif
 	? chr(7)
 	? chr(7)
 	? chr(7)
enddo
CLEAR WINDOW PIEZAS
return


procedure saldos
SET ESCAPE OFF
define window saldos from 08,2 to 21,77 panel
on key
wait window "Realizando Consulta" nowait
mtmp = "c:\access\inventa\"+sys(3)+".dbf"
select select(1)
create dbf (mtmp) (clave c(5), ;
				nombre c(40), ;
				orden n(8,0), ;
				leido n(8,0), ;
				falta n(8,0))
use
use (mtmp) alias saldos
use &traspa index &traspa exclusive alias mvtos in select(1)
use \access\inventa\ordensur order ordenesti in select(1)
select ordensur
seek str(morden,10)
do while orden = str(morden,10) and !eof()
	store cveestilo to mestilo
   	select mvtos
   	seek left(mestilo,5)
   	mleido = canp1+canp2+canp3+canp4+canp5+canp6
   	select estilos
   	seek mestilo
	
   	select ordensur
	
   	sum cant1+cant2+cant3+cant4+cant5+cant6, ;
   		surt1+surt2+surt3+surt4+surt5+surt6 to tot, sal ;
   		while orden = str(morden,10) and cveestilo = mestilo
	
   	sal = sal +mleido
   	if (sal) < (tot)
     	sele saldos
     	append blank
     	replace clave  with mestilo, ;
     			nombre with estilos.descripcio, ;
     			orden  with tot, ;
     			leido  with sal, ;
     			falta  with (tot - sal)
   	endif
   	select ordensur
enddo
use       
select mvtos
use
select saldos
msg="Oprima ESC para salir de la consulta"
@ 22,15 say space(63)
@ 22,15 get msg
clear gets
if reccount() > 0
  	go top
  	brow field clave,nombre :H="Descripcion",orden :P="999,999", ;
  				leido :P="999,999",falta :H="Faltante" :P="999,999" ;
  				noedit window saldos color scheme 10
else
  	wait window "No hay pendientes por Empacar ®Enter para continuar¯"
endif
use
@ 22,15 say space(63)
release window saldos
dele file &mtmp
ON KEY LABEL F1 DO saldos
return








*******************************************
*      FORMATO DE LETRA MAS GRANDE        *
*******************************************
PROCEDURE ANTERIOR
 	ML1=ALLTRIM(STR(MEMBARQUE,6,0))
 	ML1= REPLICATE('0',6-LEN(ML1))+ML1
 	ML2=ALLTRIM(STR(MBULTO,5,0))
 	ML2= REPLICATE('0',5-LEN(ML2))+ML2
 	=fputs(mheader1,'US')
 	=fputs(mheader1,'N')
 	=fputs(mheader1,'Q810,24')
 	=fputs(mheader1,'q816')
 	=fputs(mheader1,'S2')
 	=fputs(mheader1,'D8')
 	=fputs(mheader1,'ZT')
 	=fputs(mheader1,'A812,0,1,4,4,2,N,"&ma1"')
 	=fputs(mheader1,'A451,6,1,4,4,2,N,"&ma6"')
 	=fputs(mheader1,'A544,6,1,4,4,2,N,"&ma5"')
 	=fputs(mheader1,'A363,8,1,4,4,2,N,"&ma3"')
 	=fputs(mheader1,'A637,8,1,4,4,2,N,"&ma2"')
 	=fputs(mheader1,'A727,8,1,4,4,2,N,"&ma4"')
 	=fputs(mheader1,'A158,12,1,5,1,1,N,"CAJA     :"')
 	=fputs(mheader1,'A270,12,1,5,1,1,N,"EMBARQUE :"')
 	=fputs(mheader1,'A80,0,1,4,3,1,N,"&mrazon"')
 	=fputs(mheader1,'A199,398,1,5,2,1,N,"&ML2"')
 	=fputs(mheader1,'A290,398,1,5,2,1,N,"&ML1"')
 	=fputs(mheader1,'B210,850,0,UA0,4,4,122,B,"&ML1&ML2"')
 	=fputs(mheader1,'P1')
 	IF MCNT > 6
  		=fputs(mheader2,'US')
  		=fputs(mheader2,'N')
  		=fputs(mheader2,'Q810,24')
  		=fputs(mheader2,'q816')
  		=fputs(mheader2,'S2')
  		=fputs(mheader2,'D8')
  		=fputs(mheader2,'ZT')
  		=fputs(mheader2,'A812,0,1,4,4,2,N,"&ma7"')
  		=fputs(mheader2,'A451,6,1,4,4,2,N,"&ma12"')
  		=fputs(mheader2,'A544,6,1,4,4,2,N,"&ma11"')
  		=fputs(mheader2,'A363,8,1,4,4,2,N,"&ma9"')
  		=fputs(mheader2,'A637,8,1,4,4,2,N,"&ma8"')
  		=fputs(mheader2,'A727,8,1,4,4,2,N,"&ma10"')
  		=fputs(mheader2,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader2,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader2,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader2,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader2,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader2,'P1')
 	ENDIF
 	IF MCNT > 12
  		=fputs(mheader3,'US')
  		=fputs(mheader3,'N')
  		=fputs(mheader3,'Q810,24')
  		=fputs(mheader3,'q816')
  		=fputs(mheader3,'S2')
  		=fputs(mheader3,'D8')
  		=fputs(mheader3,'ZT')
  		=fputs(mheader3,'A812,0,1,4,4,2,N,"&ma13"')
  		=fputs(mheader3,'A451,6,1,4,4,2,N,"&ma18"')
  		=fputs(mheader3,'A544,6,1,4,4,2,N,"&ma17"')
  		=fputs(mheader3,'A363,8,1,4,4,2,N,"&ma15"')
  		=fputs(mheader3,'A637,8,1,4,4,2,N,"&ma14"')
  		=fputs(mheader3,'A727,8,1,4,4,2,N,"&ma16"')
  		=fputs(mheader3,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader3,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader3,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader3,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader3,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader3,'P1')
 	ENDIF
 	IF MCNT > 18
  		=fputs(mheader4,'US')
  		=fputs(mheader4,'N')
  		=fputs(mheader4,'Q810,24')
  		=fputs(mheader4,'q816')
  		=fputs(mheader4,'S2')
  		=fputs(mheader4,'D8')
  		=fputs(mheader4,'ZT')
  		=fputs(mheader4,'A812,0,1,4,4,2,N,"&ma19"')
  		=fputs(mheader4,'A451,6,1,4,4,2,N,"&ma24"')
  		=fputs(mheader4,'A544,6,1,4,4,2,N,"&ma23"')
  		=fputs(mheader4,'A363,8,1,4,4,2,N,"&ma21"')
  		=fputs(mheader4,'A637,8,1,4,4,2,N,"&ma20"')
  		=fputs(mheader4,'A727,8,1,4,4,2,N,"&ma22"')
  		=fputs(mheader4,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader4,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader4,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader4,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader4,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader4,'P1')
 	ENDIF
 	IF MCNT > 24
  		=fputs(mheader5,'US')
  		=fputs(mheader5,'N')
  		=fputs(mheader5,'Q810,24')
  		=fputs(mheader5,'q816')
  		=fputs(mheader5,'S2')
  		=fputs(mheader5,'D8')
  		=fputs(mheader5,'ZT')
  		=fputs(mheader5,'A812,0,1,4,4,2,N,"&ma25"')    &&1
  		=fputs(mheader5,'A451,6,1,4,4,2,N,"&ma30"')    &&6
  		=fputs(mheader5,'A544,6,1,4,4,2,N,"&ma29"')    &&5
  		=fputs(mheader5,'A363,8,1,4,4,2,N,"&ma27"')    &&3
  		=fputs(mheader5,'A637,8,1,4,4,2,N,"&ma26"')    &&2
  		=fputs(mheader5,'A727,8,1,4,4,2,N,"&ma28"')    &&4
  		=fputs(mheader5,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader5,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader5,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader5,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader5,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader5,'P1')
 	ENDIF
 	IF MCNT > 30
  		=fputs(mheader6,'US')
  		=fputs(mheader6,'N')
  		=fputs(mheader6,'Q810,24')
  		=fputs(mheader6,'q816')
  		=fputs(mheader6,'S2')
  		=fputs(mheader6,'D8')
  		=fputs(mheader6,'ZT')
  		=fputs(mheader6,'A812,0,1,4,4,2,N,"&ma31"')       &&1
  		=fputs(mheader6,'A451,6,1,4,4,2,N,"&ma36"')       &&6
  		=fputs(mheader6,'A544,6,1,4,4,2,N,"&ma35"')       &&5
  		=fputs(mheader6,'A363,8,1,4,4,2,N,"&ma33"')       &&3
  		=fputs(mheader6,'A637,8,1,4,4,2,N,"&ma32"')       &&2
  		=fputs(mheader6,'A727,8,1,4,4,2,N,"&ma34"')       &&4
  		=fputs(mheader6,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader6,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader6,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader6,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader6,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader6,'P1')
 	ENDIF
 	IF MCNT > 36
  		=fputs(mheader7,'US')
  		=fputs(mheader7,'N')
  		=fputs(mheader7,'Q810,24')
  		=fputs(mheader7,'q816')
  		=fputs(mheader7,'S2')
  		=fputs(mheader7,'D8')
  		=fputs(mheader7,'ZT')
  		=fputs(mheader7,'A812,0,1,4,4,2,N,"&ma37"')       &&1
  		=fputs(mheader7,'A451,6,1,4,4,2,N,"&ma42"')       &&6
  		=fputs(mheader7,'A544,6,1,4,4,2,N,"&ma41"')       &&5
  		=fputs(mheader7,'A363,8,1,4,4,2,N,"&ma39"')       &&3
  		=fputs(mheader7,'A637,8,1,4,4,2,N,"&ma38"')       &&2
  		=fputs(mheader7,'A727,8,1,4,4,2,N,"&ma40"')       &&4
  		=fputs(mheader7,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader7,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader7,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader7,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader7,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader7,'P1')
 	ENDIF
 	IF MCNT > 42
  		=fputs(mheader8,'US')
  		=fputs(mheader8,'N')
  		=fputs(mheader8,'Q810,24')
  		=fputs(mheader8,'q816')
  		=fputs(mheader8,'S2')
  		=fputs(mheader8,'D8')
  		=fputs(mheader8,'ZT')
  		=fputs(mheader8,'A812,0,1,4,4,2,N,"&ma43"')       &&1
  		=fputs(mheader8,'A451,6,1,4,4,2,N,"&ma48"')       &&6
  		=fputs(mheader8,'A544,6,1,4,4,2,N,"&ma47"')       &&5
  		=fputs(mheader8,'A363,8,1,4,4,2,N,"&ma45"')       &&3
  		=fputs(mheader8,'A637,8,1,4,4,2,N,"&ma44"')       &&2
  		=fputs(mheader8,'A727,8,1,4,4,2,N,"&ma46"')       &&4
  		=fputs(mheader8,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader8,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader8,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader8,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader8,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader8,'P1')
 	ENDIF
 	IF MCNT > 48
  		=fputs(mheader9,'US')
  		=fputs(mheader9,'N')
  		=fputs(mheader9,'Q810,24')
  		=fputs(mheader9,'q816')
  		=fputs(mheader9,'S2')
  		=fputs(mheader9,'D8')
  		=fputs(mheader9,'ZT')
  		=fputs(mheader9,'A812,0,1,4,4,2,N,"&ma49"')       &&1
  		=fputs(mheader9,'A451,6,1,4,4,2,N,"&ma54"')       &&6
  		=fputs(mheader9,'A544,6,1,4,4,2,N,"&ma53"')       &&5
  		=fputs(mheader9,'A363,8,1,4,4,2,N,"&ma51"')       &&3
  		=fputs(mheader9,'A637,8,1,4,4,2,N,"&ma50"')       &&2
  		=fputs(mheader9,'A727,8,1,4,4,2,N,"&ma52"')       &&4
  		=fputs(mheader9,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader9,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader9,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader9,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader9,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader9,'P1')
 	ENDIF
 	IF MCNT > 54
  		=fputs(mheader10,'US')
  		=fputs(mheader10,'N')
  		=fputs(mheader10,'Q810,24')
  		=fputs(mheader10,'q816')
  		=fputs(mheader10,'S2')
  		=fputs(mheader10,'D8')
  		=fputs(mheader10,'ZT')
  		=fputs(mheader10,'A812,0,1,4,4,2,N,"&ma55"')       &&1
  		=fputs(mheader10,'A451,6,1,4,4,2,N,"&ma60"')       &&6
  		=fputs(mheader10,'A544,6,1,4,4,2,N,"&ma59"')       &&5
  		=fputs(mheader10,'A363,8,1,4,4,2,N,"&ma57"')       &&3
  		=fputs(mheader10,'A637,8,1,4,4,2,N,"&ma56"')       &&2
  		=fputs(mheader10,'A727,8,1,4,4,2,N,"&ma58"')       &&4
  		=fputs(mheader10,'A158,12,1,5,1,1,N,"CAJA     :"')
  		=fputs(mheader10,'A280,12,1,5,1,1,N,"EMBARQUE :"')
  		=fputs(mheader10,'A80,0,1,4,3,1,N,"&mrazon"')
  		=fputs(mheader10,'A199,398,1,5,2,1,N,"&ML2"')
  		=fputs(mheader10,'A290,398,1,5,2,1,N,"&ML1"')
  		=fputs(mheader10,'P1')
 	ENDIF
RETURN