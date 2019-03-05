using System;
using System.Web.Mvc;
using System.Web.Routing;
using Ninject;
using FileGenerator.Domain.Concrete;
using FileGenerator.Domain.Abstract;

namespace FileGenerator.WebUI.Infrastructure
{
    public class NinjectControllerFactory : DefaultControllerFactory {
        private IKernel ninjectKernel;
        public NinjectControllerFactory()
        {
            ninjectKernel = new StandardKernel();
            AddBindings();
        }
        protected override IController GetControllerInstance(RequestContext
        requestContext, Type controllerType)
        {
            return controllerType == null
            ? null
            : (IController)ninjectKernel.Get(controllerType);
        }
        private void AddBindings()
        {
            // put bindings here
            ninjectKernel.Bind<IProductRepository>().To<EFProductRepository>();
            ninjectKernel.Bind<IStructRepository>().To<EFStructRepository>();
            ninjectKernel.Bind<IFieldsRepository>().To<EFFIeldRepository>();
            ninjectKernel.Bind<IStructFieldRepository>().To<EFStructFIeldRepository>();
            ninjectKernel.Bind<ILFileRepository>().To<EFLFilesRepository>();
            ninjectKernel.Bind<IDataFieldRepository>().To<EFDataFieldRepository>();

        }
    }
}