﻿#pragma checksum "..\..\Page_LanguageSetting.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "BB098A0E6387460E51BD33C7E4E563DB9CF929974A04BA5497560835662DD9B9"
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
using W_Opera;


namespace W_Opera {
    
    
    /// <summary>
    /// Page_LanguageSetting
    /// </summary>
    public partial class Page_LanguageSetting : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 29 "..\..\Page_LanguageSetting.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.RadioButton rbLanguage_Eng;
        
        #line default
        #line hidden
        
        
        #line 30 "..\..\Page_LanguageSetting.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.RadioButton rbLanguage_Kor;
        
        #line default
        #line hidden
        
        
        #line 31 "..\..\Page_LanguageSetting.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.RadioButton rbLanguage_Vni;
        
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
            System.Uri resourceLocater = new System.Uri("/W_Opera;component/page_languagesetting.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\Page_LanguageSetting.xaml"
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
            this.rbLanguage_Eng = ((System.Windows.Controls.RadioButton)(target));
            
            #line 29 "..\..\Page_LanguageSetting.xaml"
            this.rbLanguage_Eng.Click += new System.Windows.RoutedEventHandler(this.RbLanguage_Eng_Click);
            
            #line default
            #line hidden
            return;
            case 2:
            this.rbLanguage_Kor = ((System.Windows.Controls.RadioButton)(target));
            
            #line 30 "..\..\Page_LanguageSetting.xaml"
            this.rbLanguage_Kor.Click += new System.Windows.RoutedEventHandler(this.RbLanguage_Kor_Click);
            
            #line default
            #line hidden
            return;
            case 3:
            this.rbLanguage_Vni = ((System.Windows.Controls.RadioButton)(target));
            
            #line 31 "..\..\Page_LanguageSetting.xaml"
            this.rbLanguage_Vni.Click += new System.Windows.RoutedEventHandler(this.RbLanguage_Vni_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}
