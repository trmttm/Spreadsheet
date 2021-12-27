from ..BoundaryOut import PresentersABC


class Presenters(PresentersABC):
    def __init__(self):
        self._observers = []

    def attach_to_response_model_receiver(self, observer):
        self._observers.append(observer)

    def _notify(self, response_model):
        for observer in self._observers:
            observer(response_model)
