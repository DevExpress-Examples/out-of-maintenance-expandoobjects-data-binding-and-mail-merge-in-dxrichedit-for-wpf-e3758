Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Dynamic
Imports System.Linq
Imports System.Windows
Imports System.Xml.Linq

Namespace ExpandoObject_MailMerge
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Public Sub New()
			InitializeComponent()
			AddHandler Loaded, AddressOf MainWindow_Loaded

		End Sub

		Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			richEditControl1.ApplyTemplate()


            Dim weathers As Object = GetExpandoFromXml("weather.xml")
			richEditControl1.Options.MailMerge.DataSource = weathers

			richEditControl1.LoadDocument("weather_report.rtf")

			ribbonControl1.SelectedPage = pageMailings
			richEditControl1.Options.MailMerge.ViewMergedData = True
		End Sub

        Public Shared Function GetExpandoFromXml(ByVal file As String) As IList(Of Object)
            Dim weathers = New List(Of Object)()

            Dim doc = XDocument.Load(file)
            Dim nodes = _
             From node In doc.Root.Descendants("weather") _
             Select node
            For Each n In nodes
                Dim MyData As Object = New ExpandoObject()
                MyData.LastUpdateTime = String.Format("{0:o}", DateTime.Now)
                MyData.Weather = New ExpandoObject()
                For Each child In n.Descendants()

                    Dim w = TryCast(MyData.Weather, IDictionary(Of String, Object))
                    Dim atb As XAttribute = child.Attribute("data")
                    If atb IsNot Nothing Then
                        w(child.Name.LocalName) = atb.Value
                    End If
                Next child

                weathers.Add(MyData)

            Next n
            Return weathers
        End Function

	End Class
End Namespace
