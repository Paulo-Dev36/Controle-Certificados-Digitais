using Certificados_Clientes_WL.Classes;
using System.Collections.Generic;
using System.Linq.Expressions;
using System;

namespace Certificados_Clientes_WL.Repositorios
{
    public abstract class RepositorioAbstrato<T> where T : IEntidade
    {
        public abstract void Add(T x);
        public abstract void Update(T x);
        public abstract void Remove(T x);
        public abstract IEnumerable<T> GetAll();
        public abstract IEnumerable<T> Get(Expression<Func<T, bool>> predicate);
    }
}
