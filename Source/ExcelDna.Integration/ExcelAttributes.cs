//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using JetBrains.Annotations;

namespace ExcelDna.Integration
{
    /// <summary>
    /// For user-defined functions.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelFunctionAttribute : Attribute
    {
        /// <summary>
        /// By default the name of the add-in.
        /// </summary>
        public string Category = null;
        public string Prefix = null;
        public string Name = null;
        public string Description = null;
        public string HelpTopic = null;

        private bool? isVolatile;
        public bool IsVolatile {
            get { return isVolatile ?? false; }
            set { isVolatile = value; }
        }

        private bool? isHidden;
        public bool IsHidden {
            get { return isHidden ?? false; }
            set { isHidden = value; }
        }

        private bool? isExceptionSafe;
        public bool IsExceptionSafe {
            get { return isExceptionSafe ?? false; }
            set { isExceptionSafe = value; }
        }

        private bool? isMacroType;
        public bool IsMacroType {
            get { return isMacroType ?? false; }
            set { isMacroType = value; }
        }

        private bool? isThreadSafe;
        public bool IsThreadSafe {
            get { return isThreadSafe ?? false; }
            set { isThreadSafe = value; }
        }

        private bool? isClusterSafe;
        public bool IsClusterSafe {
            get { return isClusterSafe ?? false; }
            set { isClusterSafe = value; }
        }

        private bool? explicitRegistration;
        public bool ExplicitRegistration {
            get { return explicitRegistration ?? false; }
            set { explicitRegistration = value; }
        }

        private bool? suppressOverwriteError;
        public bool SuppressOverwriteError {
            get { return suppressOverwriteError ?? false; }
            set { suppressOverwriteError = value; }
        }

        public ExcelFunctionAttribute()
        {
        }

        public ExcelFunctionAttribute(string description)
        {
            Description = description;
        }

        public void MergeGroupAttributes(ExcelFunctionAttribute ca) {
            if (Prefix == null) Prefix = ca.Prefix;
            if (Description == null) Description = ca.Description;
            if (Category == null) Category = ca.Category;
            if (HelpTopic == null) HelpTopic = ca.HelpTopic;

            if (!isVolatile.HasValue) isVolatile = ca.isVolatile;
            if (!isHidden.HasValue) isHidden = ca.isHidden;
            if (!isExceptionSafe.HasValue) isExceptionSafe = ca.isExceptionSafe;
            if (!isMacroType.HasValue) isMacroType = ca.isMacroType;
            if (!isThreadSafe.HasValue) isThreadSafe = ca.isThreadSafe;
            if (!isClusterSafe.HasValue) isClusterSafe = ca.isClusterSafe;
            if (!explicitRegistration.HasValue) explicitRegistration = ca.explicitRegistration;
            if (!suppressOverwriteError.HasValue) suppressOverwriteError = ca.suppressOverwriteError;
        }
    }

    /// <summary>
    /// For the arguments of user-defined functions.
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelArgumentAttribute : Attribute
    {
        /// <summary>
        /// Arguments of type object may receive ExcelReference.
        /// </summary>
        public bool AllowReference = false;

        public string Name = null;
        public string Description = null;

        public ExcelArgumentAttribute()
        {
        }

        public ExcelArgumentAttribute(string description)
        {
            Description = description;
        }
    }

    /// <summary>
    /// For the arguments of object handles.
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter | AttributeTargets.ReturnValue | AttributeTargets.Class | AttributeTargets.Struct, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelHandleAttribute : Attribute
    {
    }

    /// <summary>
    /// To indicate that a type will be marshalled as object handles.
    /// </summary>
    [AttributeUsage(AttributeTargets.Assembly, AllowMultiple = true)]
    [MeansImplicitUse]
    public class ExcelHandleExternalAttribute : Attribute
    {
        public ExcelHandleExternalAttribute(Type type)
        {
            Type = type;
        }

        public Type Type { get; }
    }

    /// <summary>
    /// For macro commands.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelCommandAttribute : Attribute
    {
        public string Prefix = null;
        public string Name = null;
        public string Description = null;
        public string HelpTopic = null;
        public string ShortCut = null;
        public string MenuName = null;
        public string MenuText = null;

        private bool? isExceptionSafe;
        public bool IsExceptionSafe {
            get { return isExceptionSafe ?? false; }
            set { isExceptionSafe = value; }
        }

        private bool? explicitRegistration;
        public bool ExplicitRegistration {
            get { return explicitRegistration ?? false; }
            set { explicitRegistration = value; }
        }

        private bool? suppressOverwriteError;
        public bool SuppressOverwriteError {
            get { return suppressOverwriteError ?? false; }
            set { suppressOverwriteError = value; }
        }

        [Obsolete("ExcelFunctions can be declared hidden, not ExcelCommands.")]
        public bool IsHidden = false;

        public ExcelCommandAttribute()
        {
        }

        public ExcelCommandAttribute(string description)
        {
            Description = description;
        }

        public void MergeGroupAttributes(ExcelCommandAttribute ca) {
            if (Prefix == null) Prefix = ca.Prefix;
            if (Description == null) Description = ca.Description;
            if (HelpTopic == null) HelpTopic = ca.HelpTopic;
            if (MenuName == null) MenuName = ca.MenuName;

            if (!isExceptionSafe.HasValue) isExceptionSafe = ca.isExceptionSafe;
            if (!explicitRegistration.HasValue) explicitRegistration = ca.explicitRegistration;
            if (!suppressOverwriteError.HasValue) suppressOverwriteError = ca.suppressOverwriteError;
        }
    }

    /// <summary>
    /// For user-defined parameter conversions.
    /// </summary>
	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelParameterConversionAttribute : Attribute
    {
    }

    /// <summary>
    /// For user-defined function execution handlers.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelFunctionExecutionHandlerSelectorAttribute : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelFunctionProcessorAttribute : Attribute
    {
    }
}
