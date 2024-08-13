Public Module CasosPracticos_Ejemplos_Module
    Public Sub VentanaDeTexto()
        'Declaración de variables y constantes
        Dim edad As Integer = 25
        Dim edad2 As Integer = 25
        Dim sumaEdades As Integer = edad + edad2
        Dim nombre As String = "Paco"
        Dim fecha As Date = "12/08/2024"
        Dim adulto As Boolean = True

#Region "Antiguo (MsgBox)"
        'MsgBox(prompt, [buttons], [title])
        MsgBox("Usuario: " & nombre &
            vbCrLf & "Edad: " & sumaEdades &
            vbCrLf & If(adulto, "Mayor de edad", "Menor de edad") &
            vbCrLf & "Fecha: " & fecha,
            vbOKOnly, "Título (MsgBox)")
#End Region
#Region "Nuevo (MessageBox.Show)"
        'MessageBox.Show(text, [caption], [buttons], [icon], [defaultButton], [options])
        Dim res As DialogResult
        res = MessageBox.Show("¿Deseas continuar?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If res = DialogResult.Yes Then
            MessageBox.Show("Seleccionaste Sí.")
        Else
            MessageBox.Show("Seleccionaste No.")
        End If
#End Region

    End Sub
    Public Sub Operadores()
        'Declaración de variables y constantes
        Dim valor1 As Integer = 3
        Dim valor2 As Integer = 10
        Dim resultado As Integer

        resultado = valor1 + valor2     'suma
        resultado = valor1 - valor2     'resta
        resultado = valor1 / valor2     'división
        resultado = valor1 Mod valor2   'resto / Modulo
        resultado = valor1 * valor2     'multiplicación
        resultado = valor1 ^ valor2     'exponente

        MsgBox(resultado)
    End Sub

    'NO devuelve valor
    Public Sub Subrutina_Sub()
        Console.WriteLine("¡Hola, mundo!")
    End Sub
    'SÍ devuelve valor
    Public Function Funcion_Function(a As Integer, b As Integer) As Integer
        Return a + b
    End Function

    Public Sub TryCatch()
        Try
            Dim numero1 As Integer = 10
            Dim numero2 As Integer = 0
            Dim resultado As Integer

            resultado = numero1 / numero2
            Console.WriteLine("El resultado es: " & resultado)

        Catch ex As DivideByZeroException
            Console.WriteLine("Error: No se puede dividir por cero.")
        Catch ex As Exception
            Console.WriteLine("Error Genérico")
        Finally
            Console.WriteLine("Operación finalizada.")
        End Try

        Console.ReadLine() ' Pausa para ver los resultados en consola
    End Sub
End Module

Public Class BuclesCondiciones_Ejemplos_Class
    'Métodos NO Compartidos --> requerirá una instancia de la clase
    Public Sub IfElse()
        Dim nota As Integer = 85

        If nota >= 90 Then
            Console.WriteLine("Excelente")
        ElseIf nota >= 80 Then
            Console.WriteLine("Muy bien")
        ElseIf nota >= 70 Then
            Console.WriteLine("Bien")
        ElseIf nota < 69 And nota >= 50 Then
            Console.WriteLine("Raspado")
        Else
            Console.WriteLine("Necesitas mejorar")
        End If
    End Sub
    Public Sub SelectCase_Switch()
        Dim diaSemana As Integer = 3

        Select Case diaSemana
            Case 1
                Console.WriteLine("Lunes")
            Case 2
                Console.WriteLine("Martes")
            Case 3
                Console.WriteLine("Miércoles")
            Case 4
                Console.WriteLine("Jueves")
            Case 5
                Console.WriteLine("Viernes")
            Case 6
                Console.WriteLine("Sábado")
            Case 7
                Console.WriteLine("Domingo")
            Case Else
                Console.WriteLine("Día no válido")
        End Select
    End Sub
    Public Sub For_Bucle()  'Step opcional - Por defecto a 1
        For index As Integer = 1 To 6 Step 2
            Console.WriteLine("Iteración número: " & index)
        Next
    End Sub
    Public Sub ForEach()
        Dim nombres() As String = {"Ana", "Luis", "Carlos"}

        For Each name As String In nombres
            Console.WriteLine("Hola, " & name)
        Next
    End Sub

    'Métodos Compartidos
    Public Shared Sub While_Bucle()
        Dim i As Integer = 1

        While i <= 5
            Console.WriteLine("Valor de i: " & i)
            i += 1
        End While
    End Sub
    Public Shared Sub DoWhile()
        Dim contador As Integer = 0

        Do
            contador += 1
            Console.WriteLine("Contador: " & contador)
        Loop While contador < 5
    End Sub
    Public Shared Sub DoUntil()
        Dim numero As Integer = 0

        Do
            numero += 1
            Console.WriteLine("Número: " & numero)
        Loop Until numero = 5
    End Sub
End Class

Public Module Arrays_Ejemplos
    'Colecciones de tamaño fijo que almacenan elementos del mismo tipo.
    'Se accede a ellos mediante índices.

    Public Sub Arrays_Unidimensional()
        'Inicialización Directa
        Dim frutas As String() = {"Manzana", "Naranja", "Banana"}

        'Declarar e inicializar un array de enteros con 5 elementos
        Dim numeros(4) As Integer
        numeros(0) = 1
        numeros(1) = 2
        numeros(2) = 3
        numeros(3) = 4
        numeros(4) = 5

        'Recorrer el Array
        For i As Integer = 0 To numeros.Length - 1
            Console.WriteLine(numeros(i))
        Next


        ' Declare a single-dimension array of 5 numbers.
        Dim numbers(4) As Integer
        ' Declare a single-dimension array and set its 4 values.
        Dim nums = New Integer() {1, 2, 4, 8}
        ' Change the size of an existing array to 16 elements and retain the current values.
        ReDim Preserve nums(15)
        ' Redefine the size of an existing array and reset the values.
        ReDim nums(15)
    End Sub
    Public Sub Arrays_Bidimensional()
        'Inicialización Directa
        Dim frutas As String(,) = {{"Manzana", "Roja"}, {"Naranja", "Verde"}, {"Banana", "Amarilla"}}

        'Declarar e inicializar un array bidimensional
        Dim matriz(1, 1) As Integer
        matriz(0, 0) = 1
        matriz(0, 1) = 2
        matriz(1, 0) = 3
        matriz(1, 1) = 4

        'Recorre el array bidimensional
        For i As Integer = 0 To frutas.GetLength(0) - 1
            For j As Integer = 0 To frutas.GetLength(1) - 1
                Console.Write(frutas(i, j) & " ")
            Next
            Console.WriteLine() ' Salto de línea después de cada fila
        Next
    End Sub
    'Array de Arrays (cada Array interno puede tener longitudes diferentes
    Public Sub Arrays_Jagged()
        'Crear un array jagged de enteros
        Dim jaggedArray As Integer()() = New Integer(2)() {}

        'Inicializar los arrays internos con diferentes longitudes
        jaggedArray(0) = New Integer() {1, 2, 3}
        jaggedArray(1) = New Integer() {4, 5}
        jaggedArray(2) = New Integer() {6, 7, 8, 9}

        ' Recorre el array jagged
        For i As Integer = 0 To jaggedArray.Length - 1
            Console.WriteLine("Fila " & i & ":")
            For j As Integer = 0 To jaggedArray(i).Length - 1
                Console.Write(jaggedArray(i)(j) & " ")
            Next
            Console.WriteLine() ' Salto de línea después de cada fila
        Next

    End Sub
End Module

Public Module Listas_Ejempos
    'Son colecciones dinámicas que pueden cambiar de tamaño en tiempo de ejecución
    'permiten operaciones de inserción, eliminación y acceso por índice

    Public Sub DeclareInitialize()
        ' Declarar una lista vacía de enteros
        Dim listaNumeros As New List(Of Integer)

        ' Declaración directa de cadenas
        Dim listaCadenas As New List(Of String) From {"Manzana", "Naranja", "Banana"}

        ' Declarar una lista vacía de enteros
        Dim listaEnteros As New List(Of Integer)
        ' Agregar elementos a la lista
        listaEnteros.Add(1)
        listaEnteros.Add(2)
        listaEnteros.Add(3)
    End Sub
    Public Sub InitialSize()
        ' Declarar una lista de enteros con una capacidad inicial de 10
        Dim listaEnteros10 As New List(Of Integer)(10)
    End Sub
    Public Sub ListRange()
        ' Crear una lista de enteros del 1 al 10
        Dim listaRango As New List(Of Integer)(Enumerable.Range(1, 10))
    End Sub
    Public Sub ArrayToList()
        ' Inicialización desde un array
        Dim arrayEnteros As Integer() = {1, 2, 3, 4, 5}
        Dim listaDesdeArray As New List(Of Integer)(arrayEnteros)
    End Sub
End Module

Public Module Diccionarios_Ejemplos
    'Almacenan pares clave-valor, permitiendo el acceso rápido a valores basados en claves únicas

    'Para que no falle el AccessElement
    Dim diccionario As New Dictionary(Of String, Integer) From {{"Uno", 1}, {"Dos", 2}, {"Tres", 3}}

    Public Sub DeclareInitialize()
        ' Declaración directa con valores predeterminados
        Dim diccionario As New Dictionary(Of String, Integer) From {
            {"Uno", 1},
            {"Dos", 2},
            {"Tres", 3}
        }

        ' Declarar un diccionario vacío
        Dim dictionary As New Dictionary(Of String, Integer)()
        ' Agregar elementos al diccionario
        dictionary.Add("Uno", 1)
        dictionary.Add("Dos", 2)
        dictionary.Add("Tres", 3)
    End Sub
    Public Sub LateDeclare()
        'Inicializar el diccionario y añadir elementos después, sin el uso del constructor.

        ' Declaración del diccionario
        Dim diccionario As Dictionary(Of String, Integer)
        ' Inicialización tardía
        diccionario = New Dictionary(Of String, Integer)
        diccionario("Uno") = 1
        diccionario("Dos") = 2
        diccionario("Tres") = 3
    End Sub
    Public Sub PersonalizedComparator()
        'Puedes especificar un comparador personalizado para definir cómo se deben comparar las claves

        ' Crear un diccionario que no distinga entre mayúsculas y minúsculas en las claves
        Dim comparador As StringComparer = StringComparer.OrdinalIgnoreCase
        Dim diccionario As New Dictionary(Of String, Integer)(comparador)
        ' Agregar elementos
        diccionario.Add("uno", 1)
        diccionario.Add("dos", 2)
        diccionario.Add("tres", 3)
    End Sub
    Public Sub CloneDictionary()
        'Puedes crear un nuevo diccionario a partir de otro existente, copiando sus elementos

        ' Crear un diccionario existente
        Dim diccionarioOriginal As New Dictionary(Of String, Integer) From {
            {"Uno", 1},
            {"Dos", 2}
        }
        ' Clonar el diccionario existente
        Dim diccionarioClonado As New Dictionary(Of String, Integer)(diccionarioOriginal)
        ' Agregar un nuevo elemento al diccionario clonado
        diccionarioClonado.Add("Tres", 3)
    End Sub

    Public Sub AccessElement()
        Dim valor As Integer = diccionario("Dos")  ' Valor será 2
    End Sub
    Public Sub Iteration()
        For Each par As KeyValuePair(Of String, Integer) In diccionario
            Console.WriteLine("Clave: " & par.Key & ", Valor: " & par.Value)
        Next
    End Sub
    Public Sub Manipulation()
        ' Agregar, eliminar y verificar la existencia de claves
        diccionario("Cuatro") = 4
        diccionario.Remove("Uno")
        Dim existeClave As Boolean = diccionario.ContainsKey("Tres")
    End Sub
End Module

Public Module Colas_Queues_Ejemplos
    'Siguen el principio FIFO y permiten operaciones de encolar y desencolar.

    Dim cola As New Queue(Of String)()  'Iniciaiza la Cola para que no de error DeleteElement
    Public Sub DeclareInitialize()
        Dim cola As New Queue(Of String)()
        cola.Enqueue("Primero")
        cola.Enqueue("Segundo")
        cola.Enqueue("Tercero")
    End Sub
    Public Sub DeleteElement()  'Desencolar
        ' Dequeue elimina el primer elemento de la cola
        If cola.Count > 0 Then
            Dim primero As String = cola.Dequeue()
        End If
    End Sub
End Module

Public Module Pilas_Stacks_Ejemplos
    'Siguen el principio LIFO y permiten operaciones de apilar y desapilar.

    Dim pila As New Stack(Of String)()  'Pila vacía para que no de error DeleteElement
    Public Sub DeclareInitialize()
        Dim pila As New Stack(Of String)()
        pila.Push("Primero")    'Añadir elemento
        pila.Push("Segundo")
        pila.Push("Tercero")
    End Sub
    Public Sub DeleteElement()
        ' Pop elimina el elemento superior de la pila
        If pila.Count > 0 Then
            Dim superior As String = pila.Pop()
        End If
    End Sub
End Module

Public Module Conjuntos_Sets_Ejemplos
    'Son colecciones de elementos únicos y permiten operaciones como
    'agregar, eliminar y verificar la existencia de elementos.

    Dim conjunto As New HashSet(Of String)()  'Pila vacía para que no de error DeleteElement
    Public Sub DeclareInitialize()
        Dim conjunto As New HashSet(Of String)()
        conjunto.Add("Manzana")
        conjunto.Add("Naranja")
        conjunto.Add("Banana")
    End Sub
    Public Sub DeleteElement()
        ' Comprobar existencia y eliminar elementos
        Dim contiene As Boolean = conjunto.Contains("Manzana")
        conjunto.Remove("Naranja")
    End Sub
End Module

Public Module Introduccion_Y_Cast_deDatos_Ejemplos
    Public Sub Introduccion()
        'Console.Clear() 'Limpia la consola

        Console.WriteLine("Bienvenido")
        Console.WriteLine("Introduce el primer sumando...")

        'Obtenemos el dato introducido por el usuario
        Dim sumando1 As String = Console.ReadLine

        Console.WriteLine("Introduce el primer sumando...")
        Dim sumando2 As String = Console.ReadLine

        'Cast de las Strings a Int
        Console.WriteLine("El resultado es: " & CInt(sumando1) + CInt(sumando2))

        'Para que la consola espere y no se cierre automáticamente. Y ver el resultado
        'Console.ReadKey()
    End Sub
End Module

Public Module Parametros_Ejemplos
    Public Function ByVal_Ejemplo(ByVal num As Integer) As Integer
        '(por defecto y se puede omitir)
        '   crea una nueva variable en el método,
        '   como en cualquier otro lenguaje
        Return num
    End Function
    Dim numero As Integer = 3
    Public Function ByRef_Ejemplo(ByRef num As Integer) As Integer
        'Hace referencia a la variable pasada como parámetro
        ' Es decir, num es una referencia a numero (son la misma variable)
        ' Lo que le pase a numero le pasará a num
        Return num
    End Function
    Public Function Optional_Ejemplo(ByRef num As Integer, Optional num2 As Integer = 10) As Integer
        'Indica que es parámetro es opcional,
        ' pero necesita ser declarado con un valor por defecto
        'Importante añadirlos los últimos, todos los posteriores también serán opcionales
        Return num
    End Function
    Public Function ParamArray_Ejemplo(ParamArray nums() As Integer) As Integer
        'Se usa cuando no se sabe cuántos parametros se van a usar (Array de parámetros)
        ' Así no hace falta crear el mismo método con cantidades diferentes de parámetros

        'Ejemplo que calcula la media de los parámetros introducidos
        Dim res As Integer = 0
        For i As Integer = 0 To nums.Count - 1
            res += nums(i)
        Next

        Return res / nums.Count

        'Llamada al metodo (poner fuera del metodo)
        'Dim promedio1 As Integer = ParamArray_Ejemplo(60, 42, 58, 20, 100)
    End Function
End Module


#Region "Module, Class Shared & Class Instance (Teoría)"
'Module
'   Ventajas:
'       - Simplicidad: Los métodos están disponibles globalmente
'           y no es necesario crear una instancia para usarlos.
'       - Conveniencia: Ideal para operaciones y funciones que son
'           independientes del estado y se usan a nivel Global.
'   Desventajas:
'       - No Orientado a Objetos: No soporta herencia ni polimorfismo,
'           por lo que es menos flexible que las clases.
'       - Pruebas y Encapsulación: Puede ser más difícil de testear
'           y encapsular en comparación con las clases, ya que
'           los métodos del módulo están disponibles globalmente.
'
'Class Shared:
'   Ventajas:
'       - Simplicidad: Puedes llamar a métodos estáticos sin crear instancias de la clase,
'           lo que es útil para operaciones utilitarias.
'       - Uso de Recursos: No hay necesidad de instanciar la clase,
'           lo que puede ahorrar memoria y mejorar el rendimiento en ciertos casos.
'   Desventajas:
'       - Sin Estado: Los métodos Shared no pueden acceder a datos de instancia,
'           por lo que no pueden mantener el estado entre llamadas.
'       - Herencia Limitada: No se pueden sobrescribir en clases derivadas,
'           lo que limita su flexibilidad en un diseño orientado a objetos.
'
'Class Instance
'   Ventajas:
'       - Encapsulación y Estado: Permite que las instancias de la clase
'           mantengan y gestionen su propio estado.
'       - Polimorfismo: Se pueden sobrescribir en clases derivadas,
'           lo que facilita la extensibilidad y el uso de herencia.
'   Desventajas:
'       - Requiere Instancia: Necesitas crear una instancia de la clase,lo que puede ser
'           menos eficiente si solo necesitas métodos utilitarios que no requieren un estado.
'       - Más Complejo: La necesidad de gestionar instancias puede hacer el diseño más complejo.
'
'Resumen
'   - Métodos Shared son útiles para operaciones que no dependen del estado del objeto
'       y se pueden usar de manera Global sin instanciar la clase.
'   - Métodos de Instancia son ideales cuando se necesita mantener estado y aprovechar
'       características orientadas a objetos como herencia y polimorfismo.
'   - Métodos en un Módulo ofrecen una solución simple para funciones utilitarias globales
'       pero carecen de las características orientadas a objetos.
#End Region
