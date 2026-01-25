{
  description = "ooxml - Rust library for Office Open XML formats";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";
    flake-utils.url = "github:numtide/flake-utils";
    fenix = {
      url = "github:nix-community/fenix";
      inputs.nixpkgs.follows = "nixpkgs";
    };
  };

  outputs = { self, nixpkgs, flake-utils, fenix }:
    flake-utils.lib.eachDefaultSystem (system:
      let
        pkgs = nixpkgs.legacyPackages.${system};
        fenixPkgs = fenix.packages.${system};
        # Nightly toolchain for fuzzing
        nightlyToolchain = fenixPkgs.latest.withComponents [
          "cargo"
          "rustc"
          "rust-src"
          "llvm-tools-preview"
        ];
      in
      {
        devShells.default = pkgs.mkShell rec {
          buildInputs = with pkgs; [
            stdenv.cc.cc
            # Rust toolchain
            rustc
            cargo
            rust-analyzer
            clippy
            rustfmt
            # Fast linker for incremental builds
            mold
            clang
            # JS tooling for docs
            bun
            # Utilities
            unzip
          ];
          LD_LIBRARY_PATH = "${pkgs.lib.makeLibraryPath buildInputs}:$LD_LIBRARY_PATH";
        };

        # Fuzzing shell with nightly Rust
        # Usage: nix develop .#fuzz
        devShells.fuzz = pkgs.mkShell rec {
          buildInputs = with pkgs; [
            stdenv.cc.cc
            # Nightly Rust for fuzzing
            nightlyToolchain
            # Fuzzing tool
            cargo-fuzz
            # Fast linker
            mold
            clang
          ];
          LD_LIBRARY_PATH = "${pkgs.lib.makeLibraryPath buildInputs}:$LD_LIBRARY_PATH";
        };
      }
    );
}
