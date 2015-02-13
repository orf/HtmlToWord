import abc
import functools
import inspect
import contextlib
from ..operations import ChildlessOperation


def renders(operation):
    def _wrapper(func):
        func.renders_operation = operation

        @functools.wraps(func)
        def _inner(*args, **kwargs):
            return func(*args, **kwargs)

        if isinstance(operation, ChildlessOperation):
            pass  # ToDo: Do we need a context manager if it is a childless op? No. Handle it somehow.

        return contextlib.contextmanager(_inner)

    return _wrapper


class Renderer(abc.ABC):
    def __init__(self):
        self.render_methods = {}
        for name, method in inspect.getmembers(self, inspect.ismethod):
            if hasattr(method, "renders_operation"):
                # This functions renders an operation
                self.render_methods[method.renders_operation] = method

    def render(self, operations):
        for operation in operations:
            method = self.render_methods.get(operation.__class__, None)
            if method is None:
                raise NotImplementedError("Operation {0} not supported by this renderer".format(operation.__class__.__name__))

            with method(operation):
                self.render(operation.children)