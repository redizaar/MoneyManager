﻿#pragma checksum "..\..\SpecifiedImport.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "A793620D325AF0582C563AE64CBE62EC"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;
using WpfApp1;


namespace WpfApp1 {
    
    
    /// <summary>
    /// SpecifiedImport
    /// </summary>
    public partial class SpecifiedImport : System.Windows.Controls.Page, System.Windows.Markup.IComponentConnector {
        
        
        #line 211 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox transactionsRowTextBox;
        
        #line default
        #line hidden
        
        
        #line 212 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label currentFileLabel;
        
        #line default
        #line hidden
        
        
        #line 216 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox accountNumberCB;
        
        #line default
        #line hidden
        
        
        #line 217 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox accountNumberTextBox;
        
        #line default
        #line hidden
        
        
        #line 221 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox dateColumnTextBox;
        
        #line default
        #line hidden
        
        
        #line 225 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox priceColumnCB;
        
        #line default
        #line hidden
        
        
        #line 226 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox priceColumnTextBox_1;
        
        #line default
        #line hidden
        
        
        #line 227 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox priceColumnTextBox_2;
        
        #line default
        #line hidden
        
        
        #line 231 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox balanceColumnCB;
        
        #line default
        #line hidden
        
        
        #line 232 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox balanceColumnTextBox;
        
        #line default
        #line hidden
        
        
        #line 236 "..\..\SpecifiedImport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox commentColumnTextBox;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/WpfApp1;component/specifiedimport.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\SpecifiedImport.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.transactionsRowTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 2:
            this.currentFileLabel = ((System.Windows.Controls.Label)(target));
            return;
            case 3:
            this.accountNumberCB = ((System.Windows.Controls.ComboBox)(target));
            
            #line 216 "..\..\SpecifiedImport.xaml"
            this.accountNumberCB.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.accountNumberCB_SelectionChanged);
            
            #line default
            #line hidden
            return;
            case 4:
            this.accountNumberTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 5:
            this.dateColumnTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 6:
            this.priceColumnCB = ((System.Windows.Controls.ComboBox)(target));
            
            #line 225 "..\..\SpecifiedImport.xaml"
            this.priceColumnCB.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.priceColumnCB_SelectionChanged);
            
            #line default
            #line hidden
            return;
            case 7:
            this.priceColumnTextBox_1 = ((System.Windows.Controls.TextBox)(target));
            return;
            case 8:
            this.priceColumnTextBox_2 = ((System.Windows.Controls.TextBox)(target));
            return;
            case 9:
            this.balanceColumnCB = ((System.Windows.Controls.ComboBox)(target));
            
            #line 231 "..\..\SpecifiedImport.xaml"
            this.balanceColumnCB.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.balanceColumnCB_SelectionChanged);
            
            #line default
            #line hidden
            return;
            case 10:
            this.balanceColumnTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 11:
            this.commentColumnTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            }
            this._contentLoaded = true;
        }
    }
}

