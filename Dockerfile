# ===
# Docker file for sample reproducable builds with libreofficekit
# you can use this as a template for your own docker files that
# compile rust binaries which use libreofficekit
# ===

FROM rust:1.80.0-slim-bookworm

WORKDIR /app

# Copy cargo manifest and lock
COPY Cargo.toml .
COPY Cargo.lock .

# Create library entry point
RUN mkdir src && echo "" >src/lib.rs

# Build with no actual code (Dependency precaching)
RUN cargo build --target x86_64-unknown-linux-gnu --release

# Copy and build actual code
COPY src src
RUN touch src/lib.rs
RUN cargo build --target x86_64-unknown-linux-gnu --release
