   M S B A M L   `   `   `  ����B  >W_Opera, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:W_OperaW_Opera  5   	       W_Opera.Print   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   'localclr-namespace:W_Opera     5   	   -   5          � Title$  Print��5   	   $	��800��6   $	��750��6$    � WindowStartupLocation$ CenterScreen=�60   .��5
      �  � RowDefinitions 5   
   �� 5      $��40 �6   �� 5      .+�5   
   ��  �� Orientation$ 
Horizontal=�6   .+�5      �� -   5         ��Name$ btnPrint��5      $��18(�6'   $	��100��65   $��30��6A   $
��10,0q�6M   .��5      �� $ 
Horizontal=�6   .+�5      �� $��Image/Edit/printer.png��6   $��20��6<   $��20��6H   6U   �� 5      $��Print��6   $
��10,0q�6,   $��Center=�6:   $��14(�6\   6l   5      5      �� 5      -   5        $ btnXls��5      $	��100��6%   $��30��61   $
��10,0q�6=   .��5      �� $ 
Horizontal=�6   .+�5      �� $��Image/Edit/xls-export.png��6   $��20��6?   $��20��6K   6X   �� 5      $��Excel��6   $
��10,0q�6,   $��Center=�6:   $��14(�6\   6l   5      5      �� 5      -   5        $ 
btnListXls��5      $	��100��6)   $��30��65   $
��10,0q�6A   .��5      �� $ 
Horizontal=�6   .+�5      �� $��Image/Edit/xls-export.png��6   $��20��6?   $��20��6K   6X   �� 5      $��
List Excel��6   $
��10,0q�61   $��Center=�6?   $��14(�6a   6q   5       5!      5"      � 5#   
   $��1��6   .+�5$      Z� -   5        $
��view��5$       Z� Zoom$	 100Q�6)   66   5%      5&      5'      