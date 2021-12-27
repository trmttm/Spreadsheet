from abc import ABC
from abc import abstractmethod


class PresenterABC(ABC):
    @abstractmethod
    def attach(self, observer):
        pass

    @abstractmethod
    def present(self, response_model):
        pass
