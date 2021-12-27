class Catcher:
    def __init__(self):
        self._response_models = ()
        self._view_models = ()
        self._gateway_models = ()
        self._gateway_last_models = ()

    @property
    def response_models(self) -> tuple:
        return self._response_models

    @property
    def view_models(self) -> tuple:
        return self._view_models

    @property
    def gateway_models(self) -> tuple:
        return self._gateway_models

    @property
    def gateway_last_models(self) -> tuple:
        return self._gateway_last_models

    def add_response_model(self, response_model) -> tuple:
        response_models = list(self._response_models)
        response_models.append(response_model)
        self._response_models = tuple(response_models)
        return self._response_models

    def add_view_model(self, view_model) -> tuple:
        view_models = list(self._view_models)
        view_models.append(view_model)
        self._view_models = tuple(view_models)
        return self._view_models

    def add_gateway_model(self, gateway_model) -> tuple:
        gateway_models = list(self._gateway_models)
        gateway_models.append(gateway_model)
        self._gateway_models = tuple(gateway_models)
        return self._gateway_models

    def add_gateway_last_model(self, model, **_) -> tuple:
        models = list(self._gateway_last_models)
        models.append(model)
        self._gateway_last_models = tuple(models)
        return self._gateway_last_models
