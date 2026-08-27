"""Repository base."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Generic, TypeVar

T = TypeVar("T")


class BaseRepository(ABC, Generic[T]):
    @abstractmethod
    def list(self) -> list[T]:
        raise NotImplementedError

    @abstractmethod
    def get(self, key: str) -> T | None:
        raise NotImplementedError
