   M S B A M L   `   `   `  ����  WikidotSiteBackupTool&     WikidotSiteBackupTool.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5        � Title$  Wikidot Site Backup Tool��5      $	��350��6&   $	��525��63   $	��350��6?   $	��525��6O    � Icon$ Icon/Icon.ico��6^   .��5      � $��10q�6    � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $��* �6   �� 5	      $��* �6   �� 5
      $
��Auto �6   .+�5   
   � $��0��6    5      �� 5      $
��Auto �6    �� 5      $
��Auto �6    �� 5      $
��Auto �6    �� 5      $
��Auto �6     � ColumnDefinitions 5      �� 5      $
��Auto �6#   �� 5      $��* �6#   .+�5      �� $��0��6   $��0��6%   $��Target Site:��65   $��Center=�6I   �� 5      -   5         ��Name$ 
txtSiteUrl��5      $��0��6*   $��1��67   $��5,0,0,0q�6G   $��http://pc-dos.wikidot.com��6X   $��Center=�6y   �� 5      $��1��6   $��0��6%   $��
Page List:��65   $��0,5,0,0q�6G   $��Center=�6X   �� 5      -   5        $ txtListPageName��5      $��1��6/   $��1��6<   $��5,5,0,0q�6L   $��system:list-all-pages��6]   $��Center=�6z   �� 5      $��2��6   $��0��6%   $��Save to:��65   $��0,5,0,0q�6E   $��Center=�6V   � 5      $��2��6   $��1��6    $��Center=�60    5      �� 5      $��* �6'   �� 5      $
��Auto �6'   .+�5!      �� -   5        $ txtSavePath��5!      $��0��6/   $��5,5,0,0q�6?   $��
C:\Backup\��6P   $��Center=�6b   �� 5"      -   5         ��Name$ btnBrowseSavePath��5"      $��1��64   $��
_Browse...��6D   $��5,5,0,0q�6Y   5#      �� 5$      $��3��6   $��0��6&   $��2��66    �� Orientation$ 
Horizontal=�6J   $��0,5,0,0q�6c   .+�5%      �� -   5        $ btnStartBackup��5%      $��3��61   $��B_ackup��6>   �� 5&      -   5        $ btnLogin��5&      $��3��6+   $��_Login��68   $��5,0,0,0q�6I   5'      5(      (  "System.Windows.Controls.WebBrowser  5)   
   -   5          Name$ wbbWikidotSiteContainer��5)      $��1��66   $��0,10,0,0q�6C   �� 5*   
   -   5         ��Name$ lstLog��5*      $��2��6"   $��0,10,0,0q�6/   � 5+   
   -	   5        	 �Name$	 prgProgress��5+      $��3��6+   
 � Minimum$
 0Q�68    � Maximum$	 100Q�6D   $��10��6R   $��0,10,0,0q�6^   5,      5-      