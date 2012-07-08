# -*- coding: utf-8 -*-
#
module GoogleDrive
  debugger
  cattr_accessor :email, :password

  def self.open key
    @last_session = login (email || Settings.google_spreadsheet.email), (password || Settings.google_spreadsheet.password)
    @last_session.spreadsheet_by_key key
  end

  class SuperRow
    attr_accessor :config, :row_num, :groups

    delegate :params_h, :ws, :to => :config

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
              # config.ws[row_num, config.params_h[key]]
            end

            define_method "#{key}=" do |value|
              set key, value
              # config.ws[row_num, config.params_h[key]] = value
            end
          end
        end
      end

      after_init
    end

    def get *args
      key = args.join(':')

      raise "No such column name '#{key}'" unless params_h.has_key? key

      ws[ row_num, params_h[key] ]
    end

    def set key, value
      key = key.join(':') if key.is_a? Array

      ws[ row_num, params_h[key] ] = value
    end

    def after_init

    end
  end

  class SuperConfig
    attr_accessor :params_a, :params_h, :ws, :row_class

    delegate :save, :num_rows, :num_cols, :to => :ws

    def initialize ws, _row_class = SuperRow
      self.params_a = []
      self.params_h = {}
      self.row_class = _row_class
      self.ws = ws

      ws.rows[0].each_with_index do |v,i|
        params_a[i] = v
        next if v.blank?
        params_h[v] = i+1
      end

      after_init
    end

    def after_init
    end

    def super_row row_num
      row_class.new self, row_num
    end
  end

end
