﻿using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelDna.Integration.ObjectHandles
{
    internal class TaskObjectHandler
    {
        public static bool IsUserType(Type t)
        {
            return !AssemblyLoader.IsPrimitiveParameterType(t);
        }

        public static Type ReturnType()
        {
            return typeof(string);
        }

        public static LambdaExpression ProcessTaskObject(LambdaExpression functionLambda)
        {
            var createHandleMethod = typeof(TaskObjectHandler).GetMethod(nameof(CreateHandle), BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(functionLambda.ReturnType.GetGenericArguments()[0]);

            var newParams = functionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();
            var paramsArray = newParams.Select(p => Expression.Convert(p, typeof(object)));
            var paramsArrayExp = Expression.NewArrayInit(typeof(object), paramsArray);

            var innerLambda = Expression.Invoke(functionLambda, newParams);
            var callCreateHandle = Expression.Call(createHandleMethod, innerLambda);
            return Expression.Lambda(callCreateHandle, newParams);
        }

        private static async Task<string> CreateHandle<T>(Task<T> data)
        {
            object o = await data;
            return ObjectHandler.GetHandle(o.GetType().ToString(), o);
        }
    }
}
