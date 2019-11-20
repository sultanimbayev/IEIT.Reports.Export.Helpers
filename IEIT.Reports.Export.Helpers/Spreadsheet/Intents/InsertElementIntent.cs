using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Intents
{

    /// <summary>
    /// Класс для осуществления "намерения" для вставки нового дочернего элемента в другой элемент
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class InsertElementIntent<T> where T: OpenXmlElement
    {
        /// <summary>
        /// Родительский элемет. В этот элемент производится вставка нового элемента.
        /// </summary>
        private OpenXmlElement ParentElem;

        /// <summary>
        /// Вставляемый элемент
        /// </summary>
        private T newChild;

        /// <summary>
        /// Конструктор который инициализирует "Намерение" для вставки нового элемента в родительский
        /// </summary>
        /// <param name="parentElem">Родитель в корорый вставляется новый элемент</param>
        /// <param name="newChild">Новый, вставляемый элемент</param>
        public InsertElementIntent(OpenXmlElement parentElem, T newChild)
        {
            ParentElem = parentElem;
            this.newChild = newChild;
        }
        
        /// <summary>
        /// Делегат для вставки элемента после указанного
        /// </summary>
        /// <param name="refChild">Элемент, после которго вставляется новый элемент</param>
        /// <returns>true при удачной вставке, false в обратном случае</returns>
        private bool _After(OpenXmlElement refChild)
        {
            var inserted = ParentElem.InsertAfter(newChild, refChild);
            return inserted.Equals(newChild);
        }


        /// <summary>
        /// Делегат для вставки элемента до указанного
        /// </summary>
        /// <param name="refChild">Элемент, до которго вставляется новый элемент</param>
        /// <returns>true при удачной вставке, false в обратном случае</returns>
        private bool _Before(OpenXmlElement refChild)
        {
            var inserted = ParentElem.InsertBefore(newChild, refChild);
            return inserted.Equals(newChild);
        }

        /// <summary>
        /// Вставка элемента после первого из указанных, у которого значение не null.
        /// Хотя бы один элемент должен быть не null, иначе вставка не произойдет.
        /// </summary>
        /// <param name="refChilds">Элементы в порядке, после первого из которых требуется вставить новый элемент</param>
        /// <returns>true при удачной вставке, false в обратном случае</returns>
        public bool AfterOneOf(params OpenXmlElement[] refChilds)
        {
            return AfterOneOf(refChilds);
        }

        /// <summary>
        /// Вставка элемента после первого из указанных, у которого значение не null.
        /// Хотя бы один элемент должен быть не null, иначе вставка не произойдет.
        /// </summary>
        /// <param name="refChilds">Элементы в порядке, после первого из которых требуется вставить новый элемент</param>
        /// <param name="force">Укажите true для вставки элемента даже если все элементы в списке null</param>
        /// <returns>true при удачной вставке, false в обратном случае</returns>
        public bool AfterOneOf(IEnumerable<OpenXmlElement> refChilds, bool force = true)
        {
            bool success;
            if (success = ToOneOf(refChilds, _After) || !force) { return success; }
            ParentElem.PrependChild(newChild);
            return true;
        }

        /// <summary>
        /// Вставка элемента после первого из элементов с указанным типом.
        /// </summary>
        /// <param name="childTypes">Типы элементов в нужном порядке, после первого элемента данного типа из которых требуется вставить новый элемент</param>
        /// <returns>true при удачной вставки, false в обратном случае</returns>
        public bool AfterOneOf(params Type[] childTypes)
        {
            var refChilds = GetChildsFromTypes(childTypes);
            return AfterOneOf(refChilds, true);
        }
        
        /// <summary>
        /// Вставка элемента после первого из элементов с указанным типом.
        /// </summary>
        /// <param name="childTypes">Типы элементов в нужном порядке, после первого элемента данного типа из которых требуется вставить новый элемент</param>
        /// <returns>true при удачной вставки, false в обратном случае</returns>
        public bool AfterOneOf(IEnumerable<Type> childTypes, bool force = true)
        {
            var refChilds = GetChildsFromTypes(childTypes);
            return AfterOneOf(refChilds, force);
        }

        /// <summary>
        /// Вставка элемента до первого из указанных, у которого значение не null.
        /// Хотя бы один элемент должен быть не null, иначе вставка не произойдет.
        /// </summary>
        /// <param name="refChilds">Элементы в порядке, до первого из которых требуется вставить новый элемент</param>
        /// <returns>true при удачной вставки, false в обратном случае</returns>
        public bool BeforeOneOf(params OpenXmlElement[] refChilds)
        {
            return BeforeOneOf(refChilds);
        }

        /// <summary>
        /// Вставка элемента до первого из указанных, у которого значение не null.
        /// Хотя бы один элемент должен быть не null, иначе вставка не произойдет.
        /// </summary>
        /// <param name="refChilds">Элементы в порядке, до первого из которых требуется вставить новый элемент</param>
        /// <param name="force">Укажите true для вставки элемента даже если все элементы в списке null</param>
        /// <returns>true при удачной вставке, false в обратном случае</returns>
        public bool BeforeOneOf(IEnumerable<OpenXmlElement> refChilds, bool force = false)
        {
            bool success;
            if( success = ToOneOf(refChilds, _Before) || !force) { return success; }
            ParentElem.PrependChild(newChild);
            return true;
        }

        /// <summary>
        /// Вставка элемента до первого из элементов с указанным типом.
        /// Если не найдено ни одного элемента из указанных, вставка не произойдет.
        /// </summary>
        /// <param name="childTypes">Типы элементов в нужном порядке, до первого элемента данного типа из которых требуется вставить новый элемент</param>
        /// <returns>true при удачной вставки, false в обратном случае</returns>
        public bool BeforeOneOf(params Type[] childTypes)
        {
            var refChilds = GetChildsFromTypes(childTypes);
            return BeforeOneOf(refChilds);
        }

        /// <summary>
        /// Получить элементы с указанными типами в той же последоваетльности
        /// </summary>
        /// <param name="childTypes">Типы дочерних элементов</param>
        /// <returns>Элементы с указанными типами в той же последоваетльности</returns>
        private IOrderedEnumerable<OpenXmlElement> GetChildsFromTypes(IEnumerable<Type> childTypes)
        {
            var childTypesList = childTypes.ToList();
            var refChilds = ParentElem.Elements()
                .Where(ch => childTypes.Contains(ch.GetType()))
                .OrderBy(ch => childTypesList.FindIndex((cht => cht.Equals(ch.GetType()))));
            return refChilds;
        }

        /// <summary>
        /// Вставка элемента в каком либо отношении к другим дочерним элементам.
        /// </summary>
        /// <param name="refChilds">Дочерние элементы в отношении которых вставляется новый элемент</param>
        /// <param name="insertDeleg">Делегат который определяет с каким именно отношением к другим дочерним элементам будет вставлен новый элемент</param>
        /// <returns>true при удачной вставки, false в обратном случае</returns>
        private bool ToOneOf(IEnumerable<OpenXmlElement> refChilds, Func<OpenXmlElement, bool> insertDeleg)
        {
            foreach (var refChild in refChilds)
            {
                if (refChild == null) { continue; }
                return insertDeleg(refChild);
            }
            return false;
        }


    }
}
