   M S B A M L   `   `   `  ����B  >W_Opera, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:W_OperaW_Opera  5   	       W_Opera.Window_AttachFile   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   'localclr-namespace:W_Opera     5   	     � Title$  File Attach��5   	   $	��500��6   $	��900��6*    � WindowStartupLocation$ CenterScreen=�66    � 
ResizeMode$ NoResize=�6[    � 	Resources 5	       �� MergedDictionaries �� Source�� 5
   
    5      �� 5      $" StyleResourceDictionary.xaml?�6%   %    5      .��5      �  � RowDefinitions 5   
   �� 5      $��40 �6   �� 5      $��40 �6   �� 5      .+�5   
   �� $��0��6    �� Orientation$ 
Horizontal=�6"   $��Center=�6;   .+�5      	� -   5         	�Name$ rb_NewAttach��5      $��Tạo mới��60   $	��5,0q�6B   	 d� 	IsChecked$
	 TrueY�6n   6�   	� 5      -   5        $ rb_EditAttach��5      $��Chỉnh sửa��61   $	��5,0q�6E   6t   5      �� 5   
   -   5        
 ��Name$
 grid_ButtonEditor��5   "   $��1��6   $��0��6=   $ 
Horizontal=�6M   $
��Left=�6f   $��Center=�6�   .+�5      �� -   5        $��lvButtonTop��5         styleListViewHorizontal#��[  6*   $��2q�6[   ��5      �� 5      .R�5      �� -   5          styleButtonEdit#��[ 5   !   ��5   %   �� *BackGroundColor+  ColorLightGray_LightGreen �� 	Converter# [ .��5!      �� $ 
Horizontal=�6)   .+�5"   "   ��   styleImageEdit#��[ 6G   ��6(   �� *ImageSource+6q   �� 5#   "     styeTextBlockEdit#��[ 6K   ��6,   �� *ContentButton+6x   5$      5%      5&      5(      5)      � 5*   
   $��2��6   .+�5+      �� -   5         ��Name$ lvAttachFile��5+       �� View 5,      �� 5-      .P�55      ��  �� Width$	 250��6\    �� DisplayMemberBinding 6)   �� *Model+ �� Mode$ OneWay=���56      �� 57   "   $��
Model Code��67   6N   59      �� 5:      $ 50��6Z    6)   �� *Stt+$ OneWay=���5;      �� 5<   "   $	��STT��67   6G   5>      �� 5?      $	 550��6_    6)   �� *
FileName+$ OneWay=���5@      �� 5A   "   $��	File Name��67   6M   5C      5D      5F      5G      �� 5H   
   $��2��6   $��Center=�6"   $��Center=�6=   .+�5I       � -   5          �Name$ frameLoading��5I      $y�Visible=�6*     � NavigationUIVisibility$ Hidden=�6?   6a   5J      5K      5L      