﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace SHGraduationWarning.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("SHGraduationWarning.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; ?&gt;
        ///&lt;?xml-stylesheet type=&quot;text/xsl&quot; href=&quot;format.xsl&quot; ?&gt;
        ///&lt;ValidateRule Name=&quot;更新學期科目級別資訊&quot;&gt;
        ///	&lt;DuplicateDetection&gt;
        ///		&lt;Detector Name=&quot;更新學期科目級別資訊唯一值&quot;&gt;
        ///			&lt;Field Name=&quot;學號&quot; /&gt;
        ///			&lt;Field Name=&quot;狀態&quot; /&gt;
        ///			&lt;Field Name=&quot;學年度&quot; /&gt;
        ///			&lt;Field Name=&quot;學期&quot; /&gt;
        ///			&lt;Field Name=&quot;成績年級&quot; /&gt;
        ///			&lt;Field Name=&quot;科目名稱&quot; /&gt;
        ///			&lt;Field Name=&quot;科目級別&quot; /&gt;
        ///			&lt;Field Name=&quot;新科目名稱&quot; /&gt;
        ///			&lt;Field Name=&quot;新科目級別&quot; /&gt;
        ///		&lt;/Detector&gt;
        ///	&lt;/DuplicateDetection&gt;
        ///	&lt;FieldList&gt;
        ///
        ///		&lt;Field Required=&quot;True&quot; Name=&quot;學號&quot;&gt;
        ///			 [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string ImportUpdateSubjectLevelVal {
            get {
                return ResourceManager.GetString("ImportUpdateSubjectLevelVal", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized resource of type System.Byte[].
        /// </summary>
        internal static byte[] 學期成績與課規比對樣板 {
            get {
                object obj = ResourceManager.GetObject("學期成績與課規比對樣板", resourceCulture);
                return ((byte[])(obj));
            }
        }
    }
}