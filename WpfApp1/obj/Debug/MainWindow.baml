   M S B A M L   `   `   `  ����B  >WpfApp1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:WpfApp1WpfApp1  5   	       WpfApp1.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   'localclr-namespace:WpfApp1     5   	     � Title$  Money Manager��5   	   $��367.636��6   $	��880��60   .��5	      � $	��337��6   $	��Top=�6   ���#ECECEA.+�5
   
    � -   5          �Name$ 	MainFrame��5
      $��
150,70,0,0q�6#     � NavigationUIVisibility$ Hidden=�67   
���   �$��0q�6k   5      ]� 5   
   $��
0,70,730,0q�6   ���#ECECEA.+�5      �� $��	0,0,730,0q�6   ._�5      �� $��	0,0,570,0q�6   .+�5      �� ���#D9853B$	��150��66   $��	-90,0,0,0q�6B   ��5      �� 5       �� 
TargetType" ��6$    �� Triggers 5   "   :� 5   &    :� RoutedEvent$ 
MouseEnter��63   [�5   *   �� 5   .   .��5   2   �� .��5   6   t�  �� TargetProperty$ Margin�6I    t� From$ 	-90,0,0,0X�6l   	 t� To$ 0,0,0,0X�6}   	 g� DecelerationRatio$	 .4Q�6�   
 g� Duration$
 0:0:1D�6�   5   3   5   /   5   '   :� 5   &   $ 
MouseLeave��63   [�5   *   �� 5   .   .��5   2   �� .��5    6   t� $ Margin�6I   $ 0,0,0,0X�6l   $ 	-90,0,0,0X�6{   $	 .9Q�6�   $
 0:0:1D�6�   5!   3   5"   /   5$   '   5&      .+�5(      ]� -   5         ]�Name$ 
importDock��5(   5   $��0,0,0,0q�6$   .+�5)      �� -   5         ��Name$ ImportButton��5)   >   $��50��6&   $	��150��62   $��Import��6T   ��6e   ��  �� Path$ ImportPushed���5*   "   �� 5+   &   .R�5,   *   �� .+�5-   .   �� $��attachIcon2.png��64   $��30��6M    �� Right$ 15��6Y   
 �� Top$ 8��6k   �� 5.   .    �� Left$ 20��68   $ 13��6I   $��Import��6Y   $��Bookman Old Style)�6g   $��14(�6�   
�������$��0,5,0,0q�6�   5/   +   50   '   52      53      ]� 54      -   5        $ 	tableDock��54   5   $��0,0,0,0q�6$   .+�55      �� -   5        $ TableButton��55   =   $��50��6%   $	��150��61   $��Table��6R   ��6b   �� $ TablePushed���56   "   �� 57   &   .R�58   *   �� .+�59   .   �� $��tableIcon8.png��64   $��30��6L   $ 15��6X   $ 8��6j   �� 5:   .   $ 20��68   $ 13��6I   $��Table��6Y   $��Bookman Old Style)�6f   $��14(�6�   
�������$��0,5,0,0q�6�   5;   +   5<   '   5>      5?      ]� 5@      $��0,0,0,0q�6$   .+�5A      �� -   5        $ 
ExitButton��5A   =   $��50��6%   $	��150��61   $
��Exit��6Q   ��6`   �� $ 
ExitPushed���5B   "   �� 5C   &   .R�5D   *   �� .+�5E   .   �� $��exitIcon5.png��64   $��30��6K   $ 15��6W   $ 8��6i   �� 5F   .   $ 20��68   $ 13��6I   $
��Exit��6Y   $��Bookman Old Style)�6e   $��14(�6�   
�������$��0,5,0,0q�6�   5G   +   5H   '   5J      5K      5L      5M      5N      5O      �� 5P   
   
���   �$��1q�6%   $
��Left=�69   $��70��6T   $	��Top=�6`   $	��150��6x   ���#2c3e50._�5Q      �� .+�5R      ]� .+�5S      �� $��	Financial��6   $��15(�60   $��40,3,0,0q�6>   
�������$��Bookman Old Style)�6c   5T      �� 5U      $��History Importer��6   $��15(�63   $��10,0,0,0q�6A   
�������$��Bookman Old Style)�6f   5V      5W      �� 5X   
   -   5         ��Name$ importMenuTop��5X      
���   �$��0,0,0,1q�6<   $
��Left=�6V   $��70��6q   $��	150,0,0,0q�6}   $	��Top=�6�   $	��639��6�   ���#74AFAD._�5Y      ]� $��0,0,0,0q�6   .+�5Z      �� .+�5[      �� $��attachIcon2.png��6   $��70��65   $��
10,5,635,0q�6@   $��55��6T   �� 5\      $��Import menu��6   $��	80,20,0,0q�62   $��20(�6E   $��Arial)�6S   
�������6f   5]      5^      5_      �� 5`   
   -   5        $ tableMenuTop��5`      
���   �$��0,0,0,1q�6;   $
��Left=�6U   $��70��6p   $��	150,0,0,0q�6|   $	��Top=�6�   $	��639��6�   ���#74AFAD._�5a      ]� $��0,0,0,0q�6   .+�5b      �� .+�5c      �� $��tableIcon8.png��6   $��70��64   $��
10,5,635,0q�6?   $��55��6S   �� 5d      $��
Table menu��6   $��	80,20,0,0q�61   $��20(�6D   $��Arial)�6R   
�������6e   5e      5f      5g      �� 5h   
   -	   5        $ userNotification��5h      
���   �$��1,0,0,1q�6?   $
��Left=�6Y   $��70��6t   $��	785,0,0,0q�6�   $	��Top=�6�   $	��102��6�   ���#74AFAD._�5i      �� $��
-21,0,35,0q�6   .+�5j      �� .+�5k      �� $��exclamIcon.png��6   $��20��64   $��75,0,0,0q�6@   �� 5l      $��alarmIcon3.png��6   $��70��64   $��30,5,0,0q�6?   $��45��6Q   5m      5n      5o      �� 5p   
   -
   5         ��Name$ currentUserLabel��5p      $
��User��6*   $��800,40,0,267q�69   $
��Bold"�6O   5r      5s      