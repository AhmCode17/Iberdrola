   M S B A M L   `   `   `  ���� M  IResumenesIBerdrola, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:ResumenesIBerdrolaResumenesIBerdrola  #    ResumenesIBerdrola.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     2local clr-namespace:ResumenesIBerdrola       � Title$  	Iberdrola��$	��420��$	��716�� � Icon$ /Assets/maini.png�� � WindowStartupLocation$ CenterScreen=�.��� $��10q� � ColumnDefinitions �� �� �� �� �� ��  � RowDefinitions �� $	��100 ��� $��40 ��� $��40 ��� $��40 ��� $��40 ��� $��40 ��� $��30 ��� $��40 �.+��� -    ��Name$ image��$��3��$
��Left=�$��78��$	��Top=�$	��254�� J� RenderTransformOrigin$ -1.244,0.054&�$��Assets/iberdrola.png���� -    ��Name$ label��$��3��$��3��$��RESUMENES IBERDROLA��$
��Left=�$	��Top=�$��38��$	��300��$��20(�$
��Bold"�$��Arial Rounded MT Bold)��� -    ��Name$ txtRuta��$��1��$��4��$
��Left=�$��35��	 �� TextWrapping$
	 Wrap=�$	��Top=�$	��450���� -   
 ��Name$
 btnSeleccionar��$
��Left=�$��4��$��1��$��2��$	��Top=�$	��138��$��32��.����  �� Orientation$ 
Horizontal=�.+��� $��Assets/iconarchive.png��$��22���� $��5q�$��Seleccionar carpeta��$��Center=��� -   $��lstFiles��$��2��$��4��$��4��$	��450��$
��Left=��� -   $��lblFiles��$��6��$��3��$�� ��$
��Left=�$	��Top=�� -   $��pbStatus��$��7��$��4��$
��Left=� � IsIndeterminate$
 true��$y�Hidden=�$	��450��$��20���� -    ��Name$ lblFilesSuccess��$��7��$��4��$�� ��$��Center=�$��Center=��� -	   $
 btnProcesar��$
��Left=�$��4��$��2��$��2��$	��Top=�$	��138��$��32��.���� $ 
Horizontal=�.+��� $��Assets/iconProcess.png��$��22���� $��
  Procesar��$��Center=��� -
   $
 btnSalir��$
��Left=�$��4��$��5��$��2��$	��Top=�$	��138��$��32��.���� $ 
Horizontal=�.+��� $��Assets/iconClose.png��$��22���� $��  Salir��$��Center=��� -   $
 btnLog��$
��Left=�$��4��$��3��$��2��$	��Top=�$	��138��$��32��.���� $ 
Horizontal=�.+��� $��Assets/iconClose.png��$��22���� $�� Ver log de errores��$��Center=��� -    ��Name$ chkReemplazar��$��6��$��3��$B��<Reemplazar información si la central y el periodo ya existe��