   M S B A M L   `   `   `  ����B  >W_Opera, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:W_OperaW_Opera  5   	       W_Opera.Page_RejectSample   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   'localclr-namespace:W_Opera     5   	     � Title$  Reject Sample��5   	   $	��600��6   $
��1100��6,    � WindowStartupLocation$ CenterScreen=�69    � 	Resources 5	       �� MergedDictionaries �� Source�� 5
   
    5      �� 5      $" StyleResourceDictionary.xaml?�6%   %    5      .��5      �  � RowDefinitions 5   
   �� 5      $��40 �6   �� 5      .+�5   
   � $��0��6    � ColumnDefinitions 5      �� 5      �� 5      $	��150 �6#   .+�5      �� $��Center=�6   .+�5      �� -   5         ��Name$ 
txt_Reject��5      $��30��6.   $��Center=�6:   $��5q�6\   6i   5      �� 5      $��1��6   $��Center=�6)   .+�5      �� -   5         ��Name$ 
btn_Reject��5      $��Reject��6-   $��30��6>   $	��120��6J   6q   5      5       � 5!   
   $��1��6   .+�5"      �� -   5        	 ��Name$	 lvSampleReject��5"          styleConvertColorRowListView#��[  6/   $��2q�5#      
 �� View
 5$      �� 5%      .P�5&      ��  �� DisplayMemberBinding 6)   �� *	modelcode+ �� Mode$ OneWay=���5'      �� 5(   "   $��
Model Code��67   $	��150��6L   6Z   5*      �� 5+       6)   �� *seq+$ OneWay=���5,      �� 5-   "   $	��STT��67   $��50��6E   6R   5/      �� 50       6)   �� *remark+$ OneWay=���51      �� 52   "   $
��Note��67   $	��650��6F   6T   54      �� 55       6)   �� *
insempcode+$ OneWay=���56      �� 57   "   $��ID��67   $��50��6D   6Q   59      �� 5:       6)   �� *insdt+$ OneWay=���5;      �� 5<   "   $
��Date��67   $	��150��6F   6T   5>      5?      5B      5C      5E      5F      