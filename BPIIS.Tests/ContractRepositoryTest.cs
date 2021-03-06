// <copyright file="ContractRepositoryTest.cs">Copyright ©  2018</copyright>
using System;
using Microsoft.Pex.Framework;
using Microsoft.Pex.Framework.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WindowsFormsApp1.Repository;

namespace WindowsFormsApp1.Repository.Tests
{
    /// <summary>此类包含 ContractRepository 的参数化单元测试</summary>
    [PexClass(typeof(ContractRepository))]
    [PexAllowedExceptionFromTypeUnderTest(typeof(InvalidOperationException))]
    [PexAllowedExceptionFromTypeUnderTest(typeof(ArgumentException), AcceptExceptionSubtypes = true)]
    [TestClass]
    public partial class ContractRepositoryTest
    {
        /// <summary>测试 GetName(String) 的存根</summary>
        [PexMethod]
        public string GetNameTest([PexAssumeUnderTest]ContractRepository target, string wholeText)
        {
            string result = target.GetName(wholeText);
            return result;
            // TODO: 将断言添加到 方法 ContractRepositoryTest.GetNameTest(ContractRepository, String)
        }
    }
}
