import abc
import functools
import inspect
import contextlib
from ..operations import ChildlessOperation, IgnoredOperation, Group


def renders(*operations):
    def _wrapper(func):
        func.renders_operations = operations

        @functools.wraps(func)
        def _inner(*args, **kwargs):
            return func(*args, **kwargs)

        if any(isinstance(op, ChildlessOperation) for op in operations):
            if not all(isinstance(op, ChildlessOperation) for op in operations):
                raise Exception("Cannot mix ChildlessOperations and normal Operations")

            return func

        return contextlib.contextmanager(_inner)

    return _wrapper


class Renderer(abc.ABC):
    def __init__(self):
        self.render_methods = {}
        for name, method in inspect.getmembers(self, inspect.ismethod):
            if hasattr(method, "renders_operations"):
                for op in method.renders_operations:
                    self.render_methods[op] = method

    @renders(IgnoredOperation)
    def ignored_element(self, op):
        yield

    @renders(Group)
    def group(self, op):
        yield

    def render(self, operations):
        for operation in operations:
            method = self.render_methods.get(operation.__class__, None)
            if method is None:
                raise NotImplementedError("Operation {0} not supported by this renderer".format(operation.__class__.__name__))

            if isinstance(operation, ChildlessOperation):
                method(operation)
            else:
                with method(operation):
                    self.render(operation.children)