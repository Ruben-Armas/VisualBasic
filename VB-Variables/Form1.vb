Public Class Form1
    'Subrutina --> NO devuelve valor
    ' Se ejecuta al cargar el formulario
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Llamada a Métodos de ejemplo "CasosPracticos"

        'Al ser un "Module" no necesita instancia ni ser compartido, solo hacerle referencia
        CasosPracticos_Ejemplos_Module.VentanaDeTexto()
        CasosPracticos_Ejemplos_Module.Operadores()

        CasosPracticos_Ejemplos_Module.Subrutina_Sub()     'NO devuelve valor
        Dim res As Integer = CasosPracticos_Ejemplos_Module.Funcion_Function(5, 3)   'SÍ devuelve valor
        MessageBox.Show("El resultado de la función es: " & res)

        CasosPracticos_Ejemplos_Module.TryCatch()


        'Al NO ser una Class Shared, hay que instanciar la clase para hacer referencia a los métodos
        Dim bucles As New BuclesCondiciones_Ejemplos_Class() 'instancia de la clase
        bucles.IfElse()
        bucles.SelectCase_Switch()
        bucles.For_Bucle()
        bucles.ForEach()

        'Si los métodos son Class Shared, se les puede llamar sin Instanciar
        BuclesCondiciones_Ejemplos_Class.While_Bucle()
        BuclesCondiciones_Ejemplos_Class.DoWhile()
        BuclesCondiciones_Ejemplos_Class.DoUntil()


        Arrays_Ejemplos.Arrays_Unidimensional()
        Arrays_Ejemplos.Arrays_Bidimensional()
        Arrays_Ejemplos.Arrays_Jagged()

        Listas_Ejempos.DeclareInitialize()
        Listas_Ejempos.InitialSize()
        Listas_Ejempos.ListRange()
        Listas_Ejempos.ArrayToList()

        Diccionarios_Ejemplos.DeclareInitialize()
        Diccionarios_Ejemplos.LateDeclare()
        Diccionarios_Ejemplos.PersonalizedComparator()
        Diccionarios_Ejemplos.CloneDictionary()
        Diccionarios_Ejemplos.AccessElement()
        Diccionarios_Ejemplos.Iteration()
        Diccionarios_Ejemplos.Manipulation()

        Colas_Queues_Ejemplos.DeclareInitialize()
        Colas_Queues_Ejemplos.DeleteElement()

        Pilas_Stacks_Ejemplos.DeclareInitialize()
        Pilas_Stacks_Ejemplos.DeleteElement()

        Conjuntos_Sets_Ejemplos.DeclareInitialize()
        Conjuntos_Sets_Ejemplos.DeleteElement()

        Introduccion_Y_Cast_deDatos_Ejemplos.Introduccion()

        Dim promedio1 As Integer
        promedio1 = Parametros_Ejemplos.ByVal_Ejemplo(1)
        promedio1 = Parametros_Ejemplos.ByRef_Ejemplo(2)
        promedio1 = Parametros_Ejemplos.Optional_Ejemplo(3)
        promedio1 = Parametros_Ejemplos.Optional_Ejemplo(3, 4)
        promedio1 = Parametros_Ejemplos.ParamArray_Ejemplo(60, 42, 58, 20, 100)

    End Sub


    'Manejo de eventos, por ejemplo, clic en un botón
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MessageBox.Show("El botón ha sido clickeado.")
    End Sub
End Class
