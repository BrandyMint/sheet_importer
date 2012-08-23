# -*- coding: utf-8 -*-
#
require 'ostruct'

module GoogleDrive
  mattr_accessor :email, :password

  def self.open_spreadsheet key
    @last_session = login email, password

    if key=~/google.com/
      @last_session.spreadsheet_by_url key
    else
      @last_session.spreadsheet_by_key key
    end
  end

  class SuperRow
    attr_accessor :config, :row_num, :groups

    delegate :columns, :params_h, :ws, :to => :config

    def initialize _config, _row_num
      self.config = _config
      self.row_num = _row_num
      self.groups = {}

      config.params_h.keys.each do |key|
        if key=~/:/
          group, name = key.split ':'

          groups[group] ||= {}

          groups[group][name] = {
            :row => row_num,
            :col => config.params_h[key]
          }

        else
          class_eval do
            define_method key do
              get key
            end

            define_method "#{key}=" do |value|
              set key, value
            end
          end
        end
      end

      after_init
    end

    def get *args
      key = args.join(':')

      raise "No such column name '#{key}'" unless params_h.has_key? key

      value = ws[ row_num, params_h[key] ]

      config.opts.strip_values ? value.strip : value
    end

    def set key, value
      key = key.join(':') if key.is_a? Array

      ws[ row_num, params_h[key] ] = value
    end

    def after_init

    end
  end

  class SuperConfig
    attr_accessor :params_a, :params_h, :ws, :opts

    delegate :row_class, :first_content_row, :to => :opts

    delegate :save, :num_rows, :num_cols, :to => :ws

    def initialize _ws, _opts={}
      self.params_a = []
      self.params_h = {}
      self.ws = _ws

      self.opts = OpenStruct.new _opts.reverse_merge(
        :row_class => SuperRow,
        :first_content_row => 2,
        :strip_values => true
      )

      init_params

      after_init
    end

    def after_init
    end

    def super_row row_num
      row_class.new self, row_num
    end

    def columns
      params_h.keys.map &:to_sym
    end

    def each_content_row &block
      for row in first_content_row..num_rows
        block.call super_row(row)
      end
    end

    protected

    def init_params
      ws.rows[0].each_with_index do |v,i|
        params_a[i] = v
        next if v.blank?
        params_h[v] = i+1
      end
    end
  end

end
