#pragma checksum "..\..\DetailsWindow.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "99FBE91798C03BFAEF606BB4ED78C855C8B90B1A8D70BEC33307A454B7DC6FEA"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using Exercise08;
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


namespace Exercise08 {
    
    
    /// <summary>
    /// DetailsWindow
    /// </summary>
    public partial class DetailsWindow : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 16 "..\..\DetailsWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox naamTextBox;
        
        #line default
        #line hidden
        
        
        #line 17 "..\..\DetailsWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox voornaamTextBox;
        
        #line default
        #line hidden
        
        
        #line 18 "..\..\DetailsWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox geslachtTextBox;
        
        #line default
        #line hidden
        
        
        #line 19 "..\..\DetailsWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox adresTextBox;
        
        #line default
        #line hidden
        
        
        #line 20 "..\..\DetailsWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox telefoonTextBox;
        
        #line default
        #line hidden
        
        
        #line 21 "..\..\DetailsWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox geboortedatumTextBox;
        
        #line default
        #line hidden
        
        
        #line 22 "..\..\DetailsWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button bewarenButton;
        
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
            System.Uri resourceLocater = new System.Uri("/Exercise08;component/detailswindow.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\DetailsWindow.xaml"
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
            this.naamTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 2:
            this.voornaamTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 3:
            this.geslachtTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 4:
            this.adresTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 5:
            this.telefoonTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 6:
            this.geboortedatumTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 7:
            this.bewarenButton = ((System.Windows.Controls.Button)(target));
            
            #line 22 "..\..\DetailsWindow.xaml"
            this.bewarenButton.Click += new System.Windows.RoutedEventHandler(this.bewarenButton_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

