﻿#pragma checksum "..\..\..\Forms\frmLinter.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "79562577C5788CCC867A07AB0650E9F58B2B96A99D9E2607D8ECBCFA2D913854"
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
using VB2CS.Forms;
using VB2CS.UserControls;


namespace VB2CS.Forms {
    
    
    /// <summary>
    /// frmLinter
    /// </summary>
    public partial class frmLinter : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 11 "..\..\..\Forms\frmLinter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.GroupBox fraConfig;
        
        #line default
        #line hidden
        
        
        #line 12 "..\..\..\Forms\frmLinter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox txtResults;
        
        #line default
        #line hidden
        
        
        #line 13 "..\..\..\Forms\frmLinter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox txtVBPFile;
        
        #line default
        #line hidden
        
        
        #line 14 "..\..\..\Forms\frmLinter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox txtFile;
        
        #line default
        #line hidden
        
        
        #line 15 "..\..\..\Forms\frmLinter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button cmdClose;
        
        #line default
        #line hidden
        
        
        #line 16 "..\..\..\Forms\frmLinter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button cmdLint;
        
        #line default
        #line hidden
        
        
        #line 17 "..\..\..\Forms\frmLinter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label lblSrc;
        
        #line default
        #line hidden
        
        
        #line 18 "..\..\..\Forms\frmLinter.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label lblFile;
        
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
            System.Uri resourceLocater = new System.Uri("/VB2CS;component/forms/frmlinter.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\Forms\frmLinter.xaml"
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
            
            #line 9 "..\..\..\Forms\frmLinter.xaml"
            ((VB2CS.Forms.frmLinter)(target)).Loaded += new System.Windows.RoutedEventHandler(this.Form_Load);
            
            #line default
            #line hidden
            return;
            case 2:
            this.fraConfig = ((System.Windows.Controls.GroupBox)(target));
            return;
            case 3:
            this.txtResults = ((System.Windows.Controls.TextBox)(target));
            return;
            case 4:
            this.txtVBPFile = ((System.Windows.Controls.TextBox)(target));
            return;
            case 5:
            this.txtFile = ((System.Windows.Controls.TextBox)(target));
            return;
            case 6:
            this.cmdClose = ((System.Windows.Controls.Button)(target));
            
            #line 15 "..\..\..\Forms\frmLinter.xaml"
            this.cmdClose.Click += new System.Windows.RoutedEventHandler(this.cmdClose_Click);
            
            #line default
            #line hidden
            return;
            case 7:
            this.cmdLint = ((System.Windows.Controls.Button)(target));
            
            #line 16 "..\..\..\Forms\frmLinter.xaml"
            this.cmdLint.Click += new System.Windows.RoutedEventHandler(this.cmdLint_Click);
            
            #line default
            #line hidden
            return;
            case 8:
            this.lblSrc = ((System.Windows.Controls.Label)(target));
            return;
            case 9:
            this.lblFile = ((System.Windows.Controls.Label)(target));
            return;
            }
            this._contentLoaded = true;
        }
    }
}

