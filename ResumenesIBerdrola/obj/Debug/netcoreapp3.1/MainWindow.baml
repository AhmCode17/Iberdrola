   M S B A M L   `   `   `  ����M  IResumenesIBerdrola, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:ResumenesIBerdrolaResumenesIBerdrola  5   	   #    ResumenesIBerdrola.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   2local clr-namespace:ResumenesIBerdrola     5   	     � Title$  	Iberdrola��5   	   $	��420��6   $	��716��6(    � Icon$ /Assets/maini.png��64    � WindowStartupLocation$ CenterScreen=�6M   .��5
      � $��10q�6    � ColumnDefinitions 5   
   �� 5      6!   �� 5      6!   �� 5      6!   �� 5      6!   �� 5      6!   �� 5      6!    � RowDefinitions 5   
   �� 5      $	��100 �6   6+   �� 5      $��40 �6   6*   �� 5      $��40 �6   6*   �� 5      $��40 �6   6*   �� 5      $��40 �6   6*   �� 5      $��40 �6   6*   �� 5      $��30 �6   6*   �� 5      $��40 �6   6*   .+�5   
   �� -   5         ��Name$ image��5   $   $��3��6   $
��Left=�63   $��78��6N   $	��Top=�6Z   $	��254��6r    J� RenderTransformOrigin$ -1.244,0.054&�6~   $��Assets/iberdrola.png��6�   �� 5   
   -   5         ��Name$ label��5   4   $��3��6   $��3��6$   $��RESUMENES IBERDROLA��6C   $
��Left=�6a   $	��Top=�6|   $��38��6�   $	��300��6�   $��20(�6�   $
��Bold"�6�   $��Arial Rounded MT Bold)�6�   �� 5    
   -   5         ��Name$ txtRuta��5    3   $��1��6   $��4��6   $
��Left=�6D   $��35��6_   	 �� TextWrapping$
	 Wrap=�6k   $	��Top=�6   $	��450��6�   �� 5"   
   -   5        
 ��Name$
 btnSeleccionar��5"      $
��Left=�6*   $��4��6E   $��1��6U   $��2��6b   $	��Top=�6v   $	��138��6�   $��32��6�   .��5#      ��  �� Orientation$ 
Horizontal=�6   .+�5$      �� $��Assets/iconarchive.png��6   $��22��68   6E   �� 5%      $��5q�6   $��Seleccionar carpeta��6'   $��Center=�6B   6_   5&      5'      �� 5(   
   -   5        $
 btnSeleccionarBd��5(      $
��Left=�6,   $��4��6G   $��2��6W   $��2��6d   $	��Top=�6x   $	��138��6�   $��32��6�   .��5)      �� $ 
Horizontal=�6   .+�5*      �� $��Assets/iconarchive.png��6   $��22��68   6E   �� 5+      $��5q�6   $�� Seleccionar BD��6'   $��Center=�6>   6[   5,      5-      �� 5/   
   -   5        $��lstFiles��5/   m   $��2��6   $��4��6   $��4��63   $	��450��6D   $
��Left=�6Q   6m   �� 50   
   -   5        $��lblFiles��50   -   $��6��6   $��3��6   $�� ��6=   $
��Left=�6H   $	��Top=�6d   � 52   
   -   5        $��pbStatus��52   �   $��7��6   $��4��6#   $
��Left=�67    � IsIndeterminate$
 true��6S   $y�Hidden=�6j   $	��450��6~   $��20��6�   6�   �� 53   
   -	   5         ��Name$ lblFilesSuccess��53      $��7��6-   $��4��6:   $�� ��6N   $��Center=�6V   $��Center=�6s   �� 54   
   -
   5        $
 btnProcesar��54      $
��Left=�6'   $��4��6B   $��3��6R   $��2��6_   $	��Top=�6s   $	��138��6�   $��32��6�   .��55      �� $ 
Horizontal=�6   .+�56      �� $��Assets/iconProcess.png��6   $��22��68   6E   �� 57      $��
  Procesar��6   $��Center=�6/   6L   58      59      �� 5;   
   -   5        $
 btnSalir��5;      $
��Left=�6#   $��4��6>   $��5��6N   $��2��6[   $	��Top=�6o   $	��138��6�   $��32��6�   .��5<      �� $ 
Horizontal=�6   .+�5=      �� $��Assets/iconClose.png��6   $��22��66   6C   �� 5>      $��  Salir��6   $��Center=�6+   6H   5?      5@      �� 5A   
   -   5        $
 btnLog��5A      $
��Left=�6!   $��4��6<   $��4��6L   $��2��6Y   $	��Top=�6m   $	��138��6�   $��32��6�   .��5B      �� $ 
Horizontal=�6   .+�5C      �� $��Assets/iconClose.png��6   $��22��66   6C   �� 5D      $�� Ver log de errores��6   $��Center=�67   6T   5E      5F      �� 5G   
   -   5         ��Name$ chkReemplazar��5G      $��6��6*   $��3��67   $B��<Reemplazar información si la central y el periodo ya existe��6K   5I      5J      