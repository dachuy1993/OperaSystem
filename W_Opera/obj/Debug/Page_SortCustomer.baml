   M S B A M L   `   `   `  ����B  >W_Opera, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:W_OperaW_Opera  5   	       W_Opera.Page_SortCustomer   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   'localclr-namespace:W_Opera     5   	   -   5          � Title$  Customer Code��5   	   $	��450��6   $	��890��6,    � WindowStartupLocation$ CenterScreen=�68    � 
ResizeMode$ NoResize=�6]    � 	Resources 5	       �� MergedDictionaries �� Source�� 5
   
    5      �� 5      $" StyleResourceDictionary.xaml?�6%   6M   %    5      .��5      � .+�5   
   �� $��Visible=�6   .��5      �� -   5        $��lvCustomerPartCode��5          styleConvertColorRowListView#��[  5       �� View 5      �� 5      .P�5      ��  �� Width$	 150��6k    �� DisplayMemberBinding 6)   �� *TMSTMODEL_ModelCode+	 �� Mode$	 OneWay=���5      �� 5   "   -   5        $��
Model Code��5   7   

 � Tag$
 	DateInput��6L   6\   5      �� 5      $	 200��6k    6)   �� *TMSTMODEL_ModelName+$	 OneWay=���5      �� 5   "   -   5        $��
Model Name��5   7   $
 AllTimeInput��6L   6_   5       �� 5!      $	 100��6k    6)   �� *TMSTMODEL_ApplyDate+$	 OneWay=���5"      �� 5#   "   -   5        $��
Apply Date��5#   7   $
 
ModelCodel��6L   6]   5%      �� 5&      $	 100��6i    6)   �� *TMSTMODEL_Version+$	 OneWay=���5'      �� 5(   "   -   5        $��Version��5(   7   $
 
ModelCodel��6I   6Z   5*      �� 5+      $	 150��6r    6)   �� *TMSTMODEL_CustomerPartCode+$	 OneWay=���5,      �� 5-   "   -   5        $��Customer Part Code��5-   7   $
 CustomerCode��6T   6g   5/      �� 50      $	 150��6u    6)   �� *TMSTMODEL_CustomerPartCodeVer+$	 OneWay=���51      �� 52   "   -   5        $��Customer PartCodeVer��52   7   $
 PositionToName��6V   6k   54      �� 55      $	 100��6q    6)   �� *TMSTMODEL_CustomerPartVer+$	 OneWay=���56      �� 57   "   -	   5        $��Part Ver��57   7   $
 PositionToName��6J   6_   59      5=      5?      5@      5B      5C      