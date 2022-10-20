import datetime
import openpyxl
import csv
import os

Reservacion= dict()
Cliente= dict()
Sala= dict()
Turno = {1:"Matutino", 2:"Vespertino", 3:"Nocturno"}
encontradas = []
disponibles = []

libro = openpyxl.Workbook()
hoja = libro["Sheet"]
hoja.title = "Reporte "

#ARCHIVO CVS AL ULTIMO
#Reservacion
if os.path.isfile("Datos_Reservacion.csv") and os.path.isfile("Datos_Cliente.csv") and os.path.isfile("Datos_Sala.csv"):
  
  with open("Datos_Reservacion.csv","r", newline="") as archivo:
        lector = csv.reader(archivo)
        next(lector)
        
        for clave, nombre, apellido, evento, fecha, turno, sala  in lector:
            Reservacion[int(clave)] = [nombre,apellido,evento,fecha,int(turno),int(sala)]

  with open("Datos_Cliente.csv","r", newline="") as archivo:
        lector = csv.reader(archivo)
        next(lector)
        
        for clave, nombre, apellido  in lector:
            Cliente[int(clave)] = [nombre,apellido]

  with open("Datos_Sala.csv","r", newline="") as archivo:
        lector = csv.reader(archivo)
        next(lector)
        
        for clave, sala, cupo  in lector:
            Sala[int(clave)] = [sala,int(cupo)]
           
while True:
    print("")
    print("Menu Principal")
    print("")
    print("1. Reservaciones")
    print("2. Reportes.")
    print("3. Registrar una sala ")
    print("4. Registrar un cliente ")
    print("5. Salir")
    print("")
    opcion_menu=int(input("¿Que opcion del menu deseas? "))
    if opcion_menu >5:
        print("Opcion no valida")
        print("")
        continue   

    if opcion_menu == 1 :
        print("")
        print("Menu Reservaciones")
        print("")
        print("1. Registrar nueva reservacion")
        print("2. Modificar descripcion de una reservacion")
        print("3. Consultar disponibilidad de salas para una fecha")
        print("4. Salir y volver al menu principal")
        print("")
        opcion_menu_reservacion=int(input("¿Que opcion del menu deseas? "))
        if opcion_menu_reservacion >4:
            print("Opcion no valida")
            print("")
            continue
        if opcion_menu_reservacion == 1 :
          while True:
              print (list(Cliente.items()))
              nombre_cliente=input("Ingrese el nombre del cliente (Nombre y Apellido) : ")
              apellido=input("Ingrese los apellidos del cliente :")
              if (not [nombre_cliente,apellido] in Cliente.values()):
                print("No esta registrado, favor de registrarse")
                break
              while True:
                  nombre_evento=input("Ingresar el nombre del evento : ")
                  if nombre_evento == "":
                      print("No se puede dejar vacio")
                      continue
                  while True:
                      Fecha_actual=datetime.date.today()
                      fecha_reservacion=input("Ingrese fecha que deseas reservar (dd/mm/aaaa) : ")
                      Fecha_reserva=datetime.datetime.strptime(fecha_reservacion, "%d/%m/%Y" ).date()
                      dias_aprox=Fecha_reserva.day - Fecha_actual.day
                      if dias_aprox <= 2:
                          print("La reservacion se tiene que hacer con 2 dias de anticipacion")
                          continue
                      while True:
                          print(Turno.items())
                          turno=int(input("Ingrese el turno que desea reservar : "))
                          if ( not turno in Turno.keys()):
                              print("Revisar su opcion de turno seleccionada")
                              continue
                          while True:
                            print(list(Sala.items()))
                            sala=int(input("Ingresa el numero de la sala que desea reservar : "))
                            if not sala in Sala.keys():
                              print("La opcion no es valida")
                              continue

                            for nombre,apellido, evento, fecha, turno_reserva,sala_reserva in Reservacion.values():
                              if (fecha_reservacion == fecha) and (turno == turno_reserva) and(sala == sala_reserva) :
                                print("Fecha ocupada, Revisar las salas desocupadas para esa fecha")
                                break
                            else:
                                if Reservacion.keys():
                                    nueva_llave=(max(list(Reservacion.keys()))+1)
                                else:
                                    nueva_llave = 1
                                Reservacion[nueva_llave]= nombre_cliente,apellido,nombre_evento,fecha_reservacion,turno,sala
                                print("Registro realizado con exito")
                            break
                          break
                      break
                  break
              break
              

        if opcion_menu_reservacion == 2:
            while True:
                print(Reservacion.items())
                nueva_llave=int(input("Ingrese el folio de su registro : "))
                cambio=Reservacion.get(nueva_llave)
                if cambio== None:
                    print("Nombre del cliente no encontrado")
                else:
                    print("Datos Actuales:", {cambio[0]},{cambio[1]},{cambio[2]},{cambio[3]},{cambio[4]},{cambio[5]})
                    nombre_nuevo=input("Nuevo nombre del evento reservado : ")
                    Reservacion.update({nueva_llave:[nombre_cliente,apellido,nombre_nuevo,fecha_reservacion,turno,sala]})
                    print("Nombre del evento modificado")
                    break

        if opcion_menu_reservacion == 3 :
          #REVISAR ESTO
          fecha_solicitada = input("Ingrese la fecha del evento (dd/mm/aaaa): ")
          Fecha=datetime.datetime.strptime(fecha_solicitada, "%d/%m/%Y" )
          for clave,valor in list(Reservacion.items()):
            Fecha_reserva,turno,sala = (valor[3],valor[4],valor[5])
            if Fecha_reserva == fecha_solicitada:
              encontradas.append((sala,turno))
            reservas_ocupadas = set(encontradas)
          for sala in Sala.keys():
            for turno in Turno.keys():
              disponibles.append((sala,turno))
            combinaciones_resrvaciones_disponibles = set(disponibles)
          
          salas_turnos_disponibles = sorted(list(combinaciones_resrvaciones_disponibles - reservas_ocupadas))
          
          print("\n las opciones disponibles para rentar en esa fecha son : ")
          print(f"*Salas disponibles para rentar el {fecha_solicitada}*\n")
          print("Salas\t\t\t\tTurnos")
          for sala,turno in salas_turnos_disponibles:
            print(f"{sala},{Sala[sala]}\t\t{Turno[turno]}")

        if opcion_menu_reservacion == 4:
            continue

    if opcion_menu == 2:
        print("")
        print("Menu Reporte")
        print("")
        print("1. Reporte en pantalla de reservaciones para una fecha")
        print("2. Exportar reporte tabular en Excel")
        print("3. Salir y volver al menu principal")
        print("")
        opcion_menu_reporte=int(input("¿Que opcion del menu deseas? "))
        if opcion_menu_reporte >3:
            print("Opcion no valida")
            print("")
            continue

        if opcion_menu_reporte == 1:
            fecha_solicitada = input("Ingrese la fecha del evento (dd/mm/aaaa): ")
            print("**"*48)
            print("*" + " "*23 + f"REPORTE DE RESERVACIONES PARA EL DÍA {fecha_solicitada}" + " "*23 + "*")
            print("**"*48)
            print("{:<20} {:<20} {:<20} {:<20} {:<20}".format('SALA','NOMBRE','APELLIDO','EVENTO', 'TURNO' ))
            print("**"*48)
            for nueva_llave,[nombre_cliente,apellido,nombre_evento,Fecha_reserva,turno,sala] in Reservacion.items():
                if fecha_solicitada == Fecha_reserva:
                    print("{:<20} {:<20} {:<20} {:<20} {:<20}".format (sala,nombre_cliente,apellido,nombre_evento, turno ))
                    print("*"*40 + " FIN DEL REPORTE " + "*"*40)

        if opcion_menu_reporte == 2:
          elementos_sala=[(sala,nombre_cliente,apellido,nombre_evento,turno)]
          fecha_solicitada = input("Ingrese la fecha del evento (dd/mm/aaaa): ")
          hoja["A1"].value = f"REPORTE DE RESERVACIONES PARA EL DÍA {fecha_solicitada}"
          hoja["A2"].value = "SALA"
          hoja["B2"].value = "NOMBRE"
          hoja["C2"].value = "APELLIDO"
          hoja["D2"].value = "EVENTO"
          hoja["E2"].value = "TURNO"
          for  nueva_llave,[nombre_cliente,apellido,nombre_evento,Fecha,turno,sala] in Reservacion.items():
            if fecha_solicitada == Fecha_reserva:
              elementos_sala=[(sala,nombre_cliente,apellido,nombre_evento,turno)]
              for elemento in elementos_sala:
                hoja.append(elemento)
          libro.save("Reporte.xlsx")
          print("Libro creado exitosamente")

        if opcion_menu_reporte ==3:
            continue

    if opcion_menu == 3:
        while True:
            nombre_sala=input("Ingrese el nombre de la sala : ")
            if nombre_sala == "":
                print("No se puede omitir")
                continue
            else:
              while True:
                cupo_sala=int(input("Ingresar el cupo de la sala : "))
                if cupo_sala == 0 :
                  print("Debe de ser mayor a 0")
                  continue
                else:
                  if Sala.keys():
                      nueva_llave=(max(list(Sala.keys()))+1)
                  else:
                      nueva_llave = 1
                  Sala[nueva_llave] = nombre_sala, cupo_sala
                  print("Sala registrada con exito")
                  break
            break
                  
    if opcion_menu == 4:
        while True:
            nombre_cliente=input("Ingrese el nombre del cliente : ")
            if nombre_cliente == "":
                print("No se puede omitir")
                continue
            while True:
              apellido=input("Ingrese los apellidos del cliente : ")
              if apellido == "":
                  print("No se puede omitir")
                  continue
              else:
                  if Cliente.keys():    
                      nueva_llave=(max(list(Cliente.keys()))+1)
                  else:
                      nueva_llave = 1
                  Cliente[nueva_llave] = [nombre_cliente,apellido]
                  print("Cliente registrado con exito")
                  break
              break
            break

    if opcion_menu == 5 :
        print("Gracias, Vuelva pronto")
        #Reservacion
        with open("Datos_Reservacion.csv","w", newline="") as archivo:
            grabador = csv.writer(archivo)
            grabador.writerow(("Clave", "Nombre","Apellido","Evento","Fecha","Turno","Sala"))
            grabador.writerows([(clave, datos[0], datos[1], datos[2], datos[3], datos[4],datos[5]) for clave, datos in Reservacion.items()])

        #Cliente
        with open("Datos_Cliente.csv","w", newline="") as archivo:
            grabador = csv.writer(archivo)
            grabador.writerow(("Clave", "Nombre", "Apellido"))
            grabador.writerows([(clave, datos[0],datos[1]) for clave, datos in Cliente.items()])

        #Sala
        with open("Datos_Sala.csv","w", newline="") as archivo:
            grabador = csv.writer(archivo)
            grabador.writerow(("Clave", "Sala", "Cupo"))
            grabador.writerows([(clave, datos[0], datos[1]) for clave, datos in Sala.items()]) 
        break