# Linkar con Excel usando la libraria COM LinkaClientCOM.dll.


Esta demo muestra el funcionamiento de un cliente persistente trabajando con una macro Excel y mostrando la información resultante sobre una hoja de cálculo.

Ventana

Esta ventana recoge la información para hacer el login del LinkarClient.

Resultado

Se realizarán en secuencia una batería de operaciones y se presentarán los resultados en la hoja de calculo.


Registrar Libreria COM

Para usar la libreria COM es necesario que este registrada en el equipo que desea hacer uso de ella. El registro de la librería se hace mediante la herramienta RegAsm.exe cuya ubicación dependerá de la máquina donde se vaya a ejecutar. Un ejemplo de como registrar LinkarClientCOM de 32 bits desde la consola de comandos de Windows (cmd) es el siguiente:

C:\>C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm "C:\linkar\Clients\NET_Framework\x86\LinkarClientCOM.DLL" /codebase /tlb

Un ejemplo de como registrar LinkarClientCOM de 64 bits desde la consola de comandos de Windows (cmd) es el siguiente:

C:\>C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm "C:\linkar\Clients\NET_Framework\x64\LinkarClientCOM.DLL" /codebase /tlb

¡¡Muy Importante!! 

- Es recomendable abrir la consola en modo administrador para evitar fallos en el registro.

- Debe tener instalada la Framework 4.5 (v4.0.30319) y usar el RegAsm de dicha Framework.

- Existen dos versiones de RegAsm, para 32 y 64 bits (Framework o Framework64), debe usar la misma que la de la librería a registrar o el comando devolverá un error.
- Para usar la librería con Excel es necesario que la librería registrada sea de la misma versión que el Excel instalado.
