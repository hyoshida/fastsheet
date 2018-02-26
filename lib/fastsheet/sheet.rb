require 'ffi'

module Fastsheet
  class Sheet
    extend FFI::Library

    class << self
      def lib_path
        File.expand_path("../../../ext/fastsheet/target/release/libfastsheet.#{lib_ext}", __FILE__)
      end

      def lib_ext
        case RUBY_PLATFORM
        when /win32/
          # Windows
          'dll'
        when /darwin/
          # OS X
          'dylib'
        else
          # Linux, BSD
          'so'
        end
      end
    end

    ffi_lib lib_path

    attach_function :read, [:pointer, :string, :string], :pointer

    attr_reader :file_name,
                :rows, :header,
                :width, :height

    def initialize(file_name, sheet_name, options = {})
      # this method sets @rows, @height and @width
      read(this, file_name, sheet_name)

      @header = @rows.shift if options[:header]
    end

    def row(n)
      @rows[n]
    end

    def each_row
      if block_given?
        @rows.each { |r| yield r }
      else
        @rows.each
      end
    end

    def column(n)
      @rows.map { |r| r[n] }
    end

    def columns
      (0...@width).inject([]) do |cols, i|
        cols.push column(i)
      end
    end

    def each_column
      if block_given?
        columns.each { |c| yield c }
      else
        columns.each
      end
    end

    private

    def this
      address = object_id << 1
      FFI::Pointer.new(:pointer, address)
    end
  end
end
