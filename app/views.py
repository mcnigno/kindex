from flask import render_template
from flask_appbuilder.models.sqla.interface import SQLAInterface
from flask_appbuilder import ModelView, ModelRestApi, BaseView, expose

from . import appbuilder, db
from .helpers import index_init
"""
    Create your Model based REST API::

    class MyModelApi(ModelRestApi):
        datamodel = SQLAInterface(MyModel)

    appbuilder.add_api(MyModelApi)


    Create your Views::


    class MyModelView(ModelView):
        datamodel = SQLAInterface(MyModel)


    Next, register your Views::


    appbuilder.add_view(
        MyModelView,
        "My View",
        icon="fa-folder-open-o",
        category="My Category",
        category_icon='fa-envelope'
    )
"""

"""
    Application wide 404 error handler
"""

class MainView(BaseView):
    route_base = "/myview"

    @expose('/method1/<string:param1>')
    def method1(self, param1):
        # do something with param1
        # and return it
        return param1

    @expose('/method2/')
    def method2(self):
        # do something with param1
        # and render it
        #param1 = 'Hello %s' % (param1)
        return index_init()


@appbuilder.app.errorhandler(404)
def page_not_found(e):
    return (
        render_template(
            "404.html", base_template=appbuilder.base_template, appbuilder=appbuilder
        ),
        404,
    )

appbuilder.add_view_no_menu(MainView)

db.create_all()
#index_init()