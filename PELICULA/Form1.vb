Imports System.Data
Imports System.Data.OleDb
Public Class Form1
    Dim con As New OleDbConnection
    Private Sub conexionn()
        Try
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\CLICK COMPUTER\Desktop\EXAMEN III PARCIAL\Database101.accdb"
            con.Open()
            MsgBox("LA CONEXION SE REALIZO EXITOSAMENTE  ")

        Catch ex As Exception
            MessageBox.Show("NO SE REALIZO LA CONEXION", "BASE DE DATOS", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information)
        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conexionn()
        Panel1.Enabled = False
        Panel2.Enabled = False
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim comando As OleDbCommand = New OleDbCommand("INSERT INTO cliente(Id,nombre,apellidos,genero,fechnac,direccion,lugar,mail,Fingreso,avala)" & Chr(13) &
                                       "VALUES(TextBox1,TextBox2,TextBox5,ComboBox1,TextBox4,TextBox3,textBox6,textBox7,TextBox8,ComboBox2)", con)


            comando.Parameters.AddWithValue("@Id", TextBox1.Text)
            comando.Parameters.AddWithValue("@nombre", TextBox2.Text)
            comando.Parameters.AddWithValue("@apellidos", TextBox5.Text)
            comando.Parameters.AddWithValue("@genero", ComboBox1.Text)
            comando.Parameters.AddWithValue("@fechnac", TextBox4.Text)
            comando.Parameters.AddWithValue("@direccion", TextBox3.Text)
            comando.Parameters.AddWithValue("@lugar", TextBox6.Text)
            comando.Parameters.AddWithValue("@mail", TextBox7.Text)
            comando.Parameters.AddWithValue("@Fingreso", TextBox8.Text)
            comando.Parameters.AddWithValue("@avala", ComboBox2.Text)
            comando.ExecuteNonQuery()
            MsgBox("Guardado Correctamente", vbInformation, "Correcto")


        Catch ex As Exception
            MsgBox("Error al Guardar", vbExclamation, "Error")
        End Try
        mostrar1()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""


    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click

    End Sub
    Private Sub mostrar()
        DataGridView1.Visible = True

        Dim adaptador As New OleDbDataAdapter("select *From Empleado", con)
        Dim datos As New DataSet
        adaptador.Fill(datos)
        DataGridView1.DataSource = datos.Tables(0)
    End Sub
    Private Sub mostrar1()
        DataGridView1.Visible = True

        Dim adaptador As New OleDbDataAdapter("select *From cliente", con)
        Dim datos As New DataSet
        adaptador.Fill(datos)
        DataGridView1.DataSource = datos.Tables(0)
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        mostrar1()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim comando As OleDbCommand = New OleDbCommand("INSERT INTO Empleado(Id,nombre,apellidos,direccion,lugar,telef,mail,fing,fnac,sueldo)" & Chr(13) &
                                       "VALUES(TextBox9,TextBox10,TextBox11,ComboBox12,TextBox13,TextBox14,textBox15,textBox16,TextBox17,ComboBox18)", con)


            comando.Parameters.AddWithValue("@Id", TextBox9.Text)
            comando.Parameters.AddWithValue("@nombre", TextBox10.Text)
            comando.Parameters.AddWithValue("@apellidos", TextBox11.Text)
            comando.Parameters.AddWithValue("@direccion", TextBox12.Text)
            comando.Parameters.AddWithValue("@lugar", TextBox13.Text)
            comando.Parameters.AddWithValue("@telef", TextBox14.Text)
            comando.Parameters.AddWithValue("@mail", TextBox15.Text)
            comando.Parameters.AddWithValue("@fing", TextBox16.Text)
            comando.Parameters.AddWithValue("@fnac", TextBox17.Text)
            comando.Parameters.AddWithValue("@sueldo", TextBox18.Text)
            comando.ExecuteNonQuery()
            MsgBox("Guardado Correctamente", vbInformation, "Correcto")


        Catch ex As Exception
            MsgBox("Error al Guardar", vbExclamation, "Error")
        End Try
        mostrar()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel2.Enabled = True
        Panel1.Enabled = False
        DataGridView1.Visible = False
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Panel1.Enabled = True
        Panel2.Enabled = False
        DataGridView1.Visible = False

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        mostrar()

    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub
End Class
