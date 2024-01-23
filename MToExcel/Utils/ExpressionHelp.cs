using System;
using System.ArrayExtensions;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using MToExcel.Converter;

namespace MToExcel.Utils
{
    public class ExpressionHelp
    {
        /// <summary>
        /// 读取标题内容
        /// 只读取第一个对象的各种属性值
        /// </summary>
        /// <param name="exp"></param>
        /// <param name="list"></param>
        /// <param name="isTrim"></param>
        /// <returns></returns>
        public static string Read_Title_Content<T>(string exp,List<T> list,bool isTrim=true)
        {
            #region 前提准备 

            if(list==null||list.Count==0)
            {
                throw new Exception("集合不能为空");
            }

            if(string.IsNullOrEmpty(exp))
            {
                throw new Exception("标题不能为空");
            }

            #endregion

            string Real_EXP = $@"";

            if(isTrim)
            {
                Real_EXP =  exp.Trim();
            }
            else{
                Real_EXP = exp;
            }

            var first =  list.First();

            var props =  typeof(T).GetProperties();

            int Index = 1;
            props.ToList().ForEach(item=>{

                if(Real_EXP.Contains($@"${Index}"))
                {
                    Real_EXP.Replace($@"${Index}",item.GetValue(first).ToString());

                }

                Index++;

            });

            return Real_EXP;
        }

        /// <summary>
        /// 读取条件隐藏的信息
        /// 
        /// </summary>
        /// <param name="exp"></param>
        /// <returns></returns>
        public static bool Read_Condition_expression<T>(string exp,T obj)
        {
            if(string.IsNullOrEmpty(exp))
            {
                throw new ExpressionException("表达式异常");
            }

            string Real_EXP = exp.Trim();

            Type t =  typeof(T);

            var props =  t.GetProperties();

            //获取所有的属性，然后循环替代字符串中的变量
            int Index = 1;
            props.ToList().ForEach(item=>{

                if(exp.Contains($@"${Index}"))
                {
                    var varib =  props[Index-1];

                    Real_EXP.Replace($@"${Index}",varib.Name);
                    
                }    
                Index++;
            });

            try
            {
                var option =  ScriptOptions.Default.AddReferences(typeof(T).Assembly);

                var result =  CSharpScript.EvaluateAsync(Real_EXP,option,(T)obj);

                //这个表达式很简洁
                if(result.Result is not null or (object?)true)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch(Exception e)
            {
                throw new ExpressionException("表达式转化异常:详细信息为"+e.ToString());
            }
            
            
            
            
        }
    }
}