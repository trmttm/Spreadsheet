from abc import ABC
from abc import abstractmethod


class GatewayABC(ABC):
    def __init__(self):
        self._observers = []

    def attach(self, observer):
        self._observers.append(observer)

    def export(self, gateway_model, **kwargs) -> tuple:
        model = self.create_model(gateway_model, **kwargs)
        error_messages = []
        for observer in self._observers:
            feed_back = observer(model, **kwargs)
            if feed_back != 'success':
                error_messages.append(feed_back)
        return tuple(error_messages)

    @abstractmethod
    def create_model(self, gateway_model, **kwargs):
        pass
