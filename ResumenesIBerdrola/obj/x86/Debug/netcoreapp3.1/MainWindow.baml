   M S B A M L   `   `   `  ����M  IResumenesIBerdrola, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:ResumenesIBerdrolaResumenesIBerdrola  5   	   #    ResumenesIBerdrola.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   2local clr-namespace:ResumenesIBerdrola     5   	     � Title$  
MainWindow��5   	   $	��400��6   $	��800��6)   .��5	      � .+�5
   
   �� -   5         ��Name$ txtRuta��5
      $
��Left=�6#   $��35��6>   $��
45,110,0,0q�6J    �� TextWrapping$
 Wrap=�6^   $	��Top=�6r   $	��415��6�   �� 5   
   -   5         ��Name$ btnProcesar��5      $
��Left=�6'   $��635,110,0,0q�6B   $	��Top=�6W   $	��110��6o   $��35��6{   .��5      ��  �� Orientation$ 
Horizontal=�6   .+�5      �� $��Assets/iconProcess.png��6   $��22��68   6E   �� 5      $��
  Procesar��6   $��Center=�6/   6L   5      5      �� 5   
   -   5        $ btnSeleccionar��5      $
��Left=�6*   $��490,110,0,0q�6E   $	��Top=�6Z   $	��140��6r   $��35��6~   .��5      �� $ 
Horizontal=�6   .+�5      �� $��Assets/iconarchive.png��6   $��22��68   6E   �� 5      $��5q�6   $��Seleccionar carpeta��6'   $��Center=�6B   6_   5      5      �� 5   
   -   5        $��lstFiles��5   )   $��45,166,334,48q�6   6)   �� 5   
   -   5         ��Name$ image��5      $
��Left=�6   $��78��6:   $��	45,10,0,0q�6F   $	��Top=�6Y   $	��254��6q    J� RenderTransformOrigin$ -1.244,0.054&�6}   $��Assets/iberdrola.png��6�   �� 5   
   -   5        $ btnSalir��5      $
��Left=�6#   $��645,301,0,0q�6>   $	��Top=�6S   $	��100��6k   $��35��6w   .��5      �� $ 
Horizontal=�6   .+�5      �� $��Assets/iconClose.png��6   $��22��66   6C   �� 5      $��  Salir��6   $��Center=�6+   6H   5      5      �� 5   
   -   5         ��Name$ label��5      $��RESUMENES IBERDROLA��6   $
��Left=�6=   $��
470,25,0,0q�6X   $	��Top=�6l   $��38��6�   $	��265��6�   $��20(�6�   $
��Bold"�6�   $��Arial Rounded MT Bold)�6�   �� 5    
   -   5        $��lblFiles��5       $�� ��6    $
��Left=�6+   $��
45,336,0,0q�6F   $	��Top=�6Z   5!      5"      