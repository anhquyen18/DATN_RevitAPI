﻿#pragma checksum "..\..\..\MyWindows\ScheduleToExcelWindow.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "473608BC5D5CA07D35790E87E5E7D3242E3EAA3588728CAED854CC3C78CE9416"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using MaterialDesignThemes.Wpf;
using MaterialDesignThemes.Wpf.Converters;
using MaterialDesignThemes.Wpf.Transitions;
using RevitAPI_Quyen;
using RevitAPI_Quyen.MyUserControl;
using RevitAPI_Quyen.ViewModel;
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


namespace RevitAPI_Quyen.MyWindows {
    
    
    /// <summary>
    /// ScheduleToExcelWindow
    /// </summary>
    public partial class ScheduleToExcelWindow : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 25 "..\..\..\MyWindows\ScheduleToExcelWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal RevitAPI_Quyen.MyWindows.ScheduleToExcelWindow mainWindow;
        
        #line default
        #line hidden
        
        
        #line 78 "..\..\..\MyWindows\ScheduleToExcelWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox saveFileTb;
        
        #line default
        #line hidden
        
        
        #line 100 "..\..\..\MyWindows\ScheduleToExcelWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListView RevitListView;
        
        #line default
        #line hidden
        
        
        #line 131 "..\..\..\MyWindows\ScheduleToExcelWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button exportBt;
        
        #line default
        #line hidden
        
        
        #line 139 "..\..\..\MyWindows\ScheduleToExcelWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListView ToExcelListView;
        
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
            System.Uri resourceLocater = new System.Uri("/RevitAPI_Quyen;component/mywindows/scheduletoexcelwindow.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\MyWindows\ScheduleToExcelWindow.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal System.Delegate _CreateDelegate(System.Type delegateType, string handler) {
            return System.Delegate.CreateDelegate(delegateType, this, handler);
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
            this.mainWindow = ((RevitAPI_Quyen.MyWindows.ScheduleToExcelWindow)(target));
            return;
            case 2:
            this.saveFileTb = ((System.Windows.Controls.TextBox)(target));
            return;
            case 3:
            this.RevitListView = ((System.Windows.Controls.ListView)(target));
            
            #line 100 "..\..\..\MyWindows\ScheduleToExcelWindow.xaml"
            this.RevitListView.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.RevitList_SelectionChanged);
            
            #line default
            #line hidden
            return;
            case 4:
            this.exportBt = ((System.Windows.Controls.Button)(target));
            return;
            case 5:
            this.ToExcelListView = ((System.Windows.Controls.ListView)(target));
            
            #line 139 "..\..\..\MyWindows\ScheduleToExcelWindow.xaml"
            this.ToExcelListView.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.ToExcelList_SelectionChanged);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

