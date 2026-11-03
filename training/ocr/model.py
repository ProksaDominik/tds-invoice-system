"""CRNN text-line recognizer: CNN features -> 2-layer BiLSTM -> per-column class scores (CTC)."""

from __future__ import annotations

import torch
from torch import nn

from .charset import NUM_CLASSES

HEIGHT = 48
DOWNSAMPLE = 4  # the CNN halves the width twice: one output column per 4 input pixels


def _conv(cin: int, cout: int, kernel=3, padding=1) -> list[nn.Module]:
    return [nn.Conv2d(cin, cout, kernel, padding=padding, bias=False), nn.BatchNorm2d(cout), nn.ReLU(inplace=True)]


class CRNN(nn.Module):
    def __init__(self, num_classes: int = NUM_CLASSES, hidden: int = 256):
        super().__init__()
        self.cnn = nn.Sequential(
            *_conv(1, 64), nn.MaxPool2d(2, 2),  # 24 x W/2
            *_conv(64, 128), nn.MaxPool2d(2, 2),  # 12 x W/4
            *_conv(128, 256), *_conv(256, 256), nn.MaxPool2d((2, 1)),  # 6 x W/4
            *_conv(256, 512), *_conv(512, 512), nn.MaxPool2d((2, 1)),  # 3 x W/4
            *_conv(512, 512, kernel=(3, 1), padding=0),  # 1 x W/4
            nn.Dropout2d(0.1),
        )
        self.rnn = nn.LSTM(512, hidden, num_layers=2, bidirectional=True, batch_first=True, dropout=0.2)
        self.head = nn.Linear(2 * hidden, num_classes)

    def forward(self, images: torch.Tensor) -> torch.Tensor:
        """images: (batch, 1, 48, width), ink ~1 and paper ~0 -> logits (batch, width // 4, classes)."""
        features = self.cnn(images).squeeze(2).transpose(1, 2)
        features, _ = self.rnn(features)
        return self.head(features)


def output_lengths(widths: torch.Tensor) -> torch.Tensor:
    return torch.div(widths, DOWNSAMPLE, rounding_mode="floor")
